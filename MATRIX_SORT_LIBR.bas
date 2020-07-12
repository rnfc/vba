Attribute VB_Name = "MATRIX_SORT_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : ARRAY_SHELL_SORT_FUNC
'DESCRIPTION   : Shell Sort Function
'LIBRARY       : MATRIX
'GROUP         : SORT
'ID            : 007
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function ARRAY_SHELL_SORT_FUNC(ByRef DATA_ARR As Variant)
  
Dim hh As Long
Dim ii As Long
Dim jj As Long
Dim kk As Long

Dim SROW As Long
Dim NROWS As Long

Dim TEMP_STR As Variant

On Error GoTo ERROR_LABEL

SROW = LBound(DATA_ARR)
NROWS = UBound(DATA_ARR)
hh = NROWS - SROW + 1
kk = 1
Do
    kk = kk * 3 + 1
Loop While kk <= hh
Do
    kk = kk \ 3
    For ii = SROW + kk To NROWS
        TEMP_STR = DATA_ARR(ii)
        For jj = ii - kk To SROW Step -kk
            If DATA_ARR(jj) > TEMP_STR Then
                DATA_ARR(jj + kk) = DATA_ARR(jj)
            Else
                Exit For
            End If
        Next jj
        DATA_ARR(jj + kk) = TEMP_STR
    Next ii
Loop While kk > 1

ARRAY_SHELL_SORT_FUNC = DATA_ARR
'SROW = LBound(DATA_ARR): NROWS = UBound(DATA_ARR)
'Do
'    CHG_FLAG = False
'    For i = SROW To NROWS - 1
'        If (DATA_ARR(i) > DATA_ARR(i + 1) _
'            And SORT_OPT) Or (DATA_ARR(i) < _
'            DATA_ARR(i + 1) And Not SORT_OPT) Then
            ' These two need to be swapped
'            SWAP_VAL = DATA_ARR(i)
'            DATA_ARR(i) = DATA_ARR(i + 1)
'            DATA_ARR(i + 1) = SWAP_VAL
'            CHG_FLAG = True
'        End If
'    Next i
'Loop Until Not CHG_FLAG

'Sorts a vector into ascending numerical order
'DATA_MATRIX = DATA_RNG: NROWS = UBound(DATA_MATRIX, 1): NCOLUMNS = UBound(DATA_MATRIX, 2): ReDim TEMP_ARR(1 To NCOLUMNS)
'For i = NROWS - 1 To 1 Step -1
'    For j = 1 To i
'        If (DATA_MATRIX(j, 1) > DATA_MATRIX(j + 1, 1)) Then
'            For k = 1 To NCOLUMNS
'                TEMP_ARR(k) = DATA_MATRIX(j + 1, k)
'                DATA_MATRIX(j + 1, k) = DATA_MATRIX(j, k)
'                DATA_MATRIX(j, k) = TEMP_ARR(k)
'            Next k
'        End If
'    Next j
'Next i
Exit Function
ERROR_LABEL:
ARRAY_SHELL_SORT_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : IS_NUMERIC_ARRAY_SORTED_FUNC
'DESCRIPTION   : This examines the array DATA_ARR (which must contain all
'numeric values) and returns True if DATA_ARR is in sorted
'order or False if it is not in sorted order.

'LIBRARY       : MATRIX
'GROUP         : SORT
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function IS_NUMERIC_ARRAY_SORTED_FUNC(ByRef DATA_ARR As Variant) As Boolean

Dim i As Long
Dim j As Long

Dim SROW As Long
Dim NROWS As Long

On Error GoTo ERROR_LABEL

SROW = LBound(DATA_ARR)
NROWS = UBound(DATA_ARR)

'---------------------------------------------------------------------
If IsArray(DATA_ARR) = False Then
    IS_NUMERIC_ARRAY_SORTED_FUNC = False
    Exit Function
End If
'---------------------------------------------------------------------
If NROWS - SROW + 1 = 1 Then
    ' array has one element.
    IS_NUMERIC_ARRAY_SORTED_FUNC = True
    Exit Function
End If
'---------------------------------------------------------------------
If DATA_ARR(SROW) < DATA_ARR(SROW + 1) Then
    j = 1
ElseIf DATA_ARR(SROW) = DATA_ARR(SROW + 1) Then
    j = 0
ElseIf DATA_ARR(SROW) > DATA_ARR(SROW + 1) Then
    j = -1
End If
'---------------------------------------------------------------------

'---------------------------------------------------------------------
For i = SROW To NROWS - 1
'---------------------------------------------------------------------
    If DATA_ARR(i) < DATA_ARR(i + 1) Then
'---------------------------------------------------------------------
        If j = 0 Then
            j = 1
        ElseIf j = 1 Then
            ' ok
        ElseIf j = -1 Then
            IS_NUMERIC_ARRAY_SORTED_FUNC = False
            Exit Function
        End If
'---------------------------------------------------------------------
    ElseIf DATA_ARR(i) > DATA_ARR(i + 1) Then
'---------------------------------------------------------------------
        If j = 0 Then
            j = -1
        ElseIf j = -1 Then
            ' ok
        ElseIf j = 1 Then
            IS_NUMERIC_ARRAY_SORTED_FUNC = False
            Exit Function
        End If
'---------------------------------------------------------------------
    End If
'---------------------------------------------------------------------
Next i
'---------------------------------------------------------------------

IS_NUMERIC_ARRAY_SORTED_FUNC = True

Exit Function
ERROR_LABEL:
IS_NUMERIC_ARRAY_SORTED_FUNC = False
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : IS_STRING_ARRAY_SORTED_FUNC
'DESCRIPTION   :
'This examines the array DATA_ARR (which must contain all string values) and
'returns True if DATA_ARR is in sorted order or False if it is not in sorted order.

'LIBRARY       : MATRIX
'GROUP         : SORT
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function IS_STRING_ARRAY_SORTED_FUNC(ByRef DATA_ARR As Variant, _
ByRef COMPARE_METHOD As VbCompareMethod) As Boolean

Dim i As Long
Dim j As Long

Dim SROW As Long
Dim NROWS As Long

On Error GoTo ERROR_LABEL

SROW = LBound(DATA_ARR)
NROWS = UBound(DATA_ARR)

'---------------------------------------------------------------------
If IsArray(DATA_ARR) = False Then
    IS_STRING_ARRAY_SORTED_FUNC = False
    Exit Function
End If
If NROWS - SROW + 1 = 1 Then
    ' array has one element.
    IS_STRING_ARRAY_SORTED_FUNC = True
    Exit Function
End If
'---------------------------------------------------------------------
If StrComp(DATA_ARR(SROW), DATA_ARR(SROW + 1), COMPARE_METHOD) = -1 Then
    j = -1
ElseIf StrComp(DATA_ARR(SROW), DATA_ARR(SROW + 1), COMPARE_METHOD) = 0 Then
    j = 0
ElseIf StrComp(DATA_ARR(SROW), DATA_ARR(SROW + 1), COMPARE_METHOD) = 1 Then
    j = 1
End If
'---------------------------------------------------------------------

'---------------------------------------------------------------------
For i = SROW To NROWS - 1
'---------------------------------------------------------------------
    If StrComp(DATA_ARR(i), DATA_ARR(i + 1), COMPARE_METHOD) = 1 Then
'---------------------------------------------------------------------
        If j = 0 Then
            j = 1
        ElseIf j = 1 Then
            ' ok
        ElseIf j = -1 Then
            IS_STRING_ARRAY_SORTED_FUNC = False
            Exit Function
        End If
'---------------------------------------------------------------------
    ElseIf StrComp(DATA_ARR(i), DATA_ARR(i + 1), COMPARE_METHOD) = -1 Then
'---------------------------------------------------------------------
        If j = 0 Then
            j = -1
        ElseIf j = -1 Then
            ' ok
        ElseIf j = 1 Then
            IS_STRING_ARRAY_SORTED_FUNC = False
            Exit Function
        End If
'---------------------------------------------------------------------
    End If
'---------------------------------------------------------------------
Next i
'---------------------------------------------------------------------

IS_STRING_ARRAY_SORTED_FUNC = True

Exit Function
ERROR_LABEL:
IS_STRING_ARRAY_SORTED_FUNC = False
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_SWAP_SORT_FUNC
'DESCRIPTION   : Sorting Routine with swapping algorithm DATA_RNG may be matrix
'(N x M) or vector (N) Sort is always based on the first column
'ORDER_OPT = 1 - Ascending), 0 - Descending. Note: it's simple but slow. Use
'only in non critical part

'LIBRARY       : MATRIX
'GROUP         : SORT
'ID            : 009
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************


Function MATRIX_SWAP_SORT_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal ORDER_OPT As Integer = 1)

Dim i As Long
Dim k As Long
Dim j As Long

Dim SROW As Long
Dim NROWS As Long
Dim SCOLUMN As Long
Dim NCOLUMNS As Long

Dim TEMP_VAL As Variant
Dim EXCHANGED_FLAG As Boolean
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG

SROW = LBound(DATA_MATRIX, 1)
NROWS = UBound(DATA_MATRIX, 1)
SCOLUMN = LBound(DATA_MATRIX, 2)
NCOLUMNS = UBound(DATA_MATRIX, 2)

'Sorting algortithm begin
Do
    EXCHANGED_FLAG = False
    For i = SROW To NROWS Step 2
        k = i + 1
        If k > NROWS Then Exit For
        If (DATA_MATRIX(i, SCOLUMN) > DATA_MATRIX(k, SCOLUMN) And ORDER_OPT = 1) Or _
           (DATA_MATRIX(i, SCOLUMN) < DATA_MATRIX(k, SCOLUMN) And ORDER_OPT = 0) Then
            'swap rows
            For j = SCOLUMN To NCOLUMNS
                TEMP_VAL = DATA_MATRIX(k, j)
                DATA_MATRIX(k, j) = DATA_MATRIX(i, j)
                DATA_MATRIX(i, j) = TEMP_VAL
            Next j
            EXCHANGED_FLAG = True
        End If
    Next
    If SROW = LBound(DATA_MATRIX, 1) Then
        SROW = LBound(DATA_MATRIX, 1) + 1
    Else
        SROW = LBound(DATA_MATRIX, 1)
    End If
Loop Until EXCHANGED_FLAG = False And SROW = LBound(DATA_MATRIX, 1)

MATRIX_SWAP_SORT_FUNC = DATA_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_SWAP_SORT_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_HEAPS_SORT_FUNC
'DESCRIPTION   : Sorts an array into ascending numerical order using
'the Heapsort Bibliography: Numerical Recipes in C. 1992
'LIBRARY       : MATRIX
'GROUP         : SORT
'ID            : 010
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_HEAPS_SORT_FUNC(ByRef DATA_RNG As Variant, _
Optional ByRef ACOLUMN As Long = 1, _
Optional ByRef VERSION As Integer = 1)

' DATA_RNG = array to sort
' ACOLUMN = sort column
' VERSION = 1 (ascending) ; Else (descending)

Dim i As Long
Dim j As Long
Dim k As Long

Dim ii As Long
Dim jj As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_ARR As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

If VERSION <> 1 Then: VERSION = -1 '(descending order)

DATA_MATRIX = DATA_RNG
NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

ReDim TEMP_ARR(1 To NCOLUMNS)
If (NROWS < 2) Then GoTo ERROR_LABEL

ii = CInt(NROWS / 2 + 1)
jj = NROWS

Do
    If (ii > 1) Then
        ii = ii - 1
        For k = 1 To NCOLUMNS
            TEMP_ARR(k) = DATA_MATRIX(ii, k)
        Next k
    Else
        For k = 1 To NCOLUMNS
            TEMP_ARR(k) = DATA_MATRIX(jj, k)
        Next k
        For k = 1 To NCOLUMNS
            DATA_MATRIX(jj, k) = DATA_MATRIX(1, k)
        Next k
        jj = jj - 1
        If (jj = 1) Then
            For k = 1 To NCOLUMNS
                DATA_MATRIX(1, k) = TEMP_ARR(k)
            Next k
            Exit Do
        End If
    End If
    i = ii
    j = ii + ii
    Do While (j <= jj)
        If (j < jj) Then
            If VERSION * (DATA_MATRIX(j, ACOLUMN) - _
                DATA_MATRIX(j + 1, ACOLUMN)) < 0 Then j = j + 1
        End If
        If VERSION * (TEMP_ARR(ACOLUMN) - _
           DATA_MATRIX(j, ACOLUMN)) < 0 Then
            For k = 1 To NCOLUMNS
                DATA_MATRIX(i, k) = DATA_MATRIX(j, k)
            Next k
            i = j
            j = j + j
        Else
            j = jj + 1
        End If
    Loop
    For k = 1 To NCOLUMNS
        DATA_MATRIX(i, k) = TEMP_ARR(k)
    Next k
Loop

MATRIX_HEAPS_SORT_FUNC = DATA_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_HEAPS_SORT_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_DOUBLE_SORT_FUNC

'DESCRIPTION   : Double sorts an array into ascending numerical order.
'The second sort field is column No.2 in the array. If you use a 1x1
'vector, there’s no second sort field.

'LIBRARY       : MATRIX
'GROUP         : SORT
'ID            : 011
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_DOUBLE_SORT_FUNC(ByRef DATA_RNG As Variant)

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim SROW As Long
Dim NROWS As Long

Dim NCOLUMNS As Long

Dim DATA_MATRIX As Variant

Dim ATEMP_VECTOR As Variant
Dim BTEMP_VECTOR As Variant
Dim CTEMP_VECTOR As Variant
    
Dim COMP_FLAG As Boolean
Dim SWITCH_FLAG As Boolean

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
SROW = LBound(DATA_MATRIX, 1)
NROWS = UBound(DATA_MATRIX, 1)

NCOLUMNS = UBound(DATA_MATRIX, 2)

For i = SROW To NROWS - 1
    For j = i + 1 To NROWS
        COMP_FLAG = False
        For k = 1 To 2
          If DATA_MATRIX(i, k) > DATA_MATRIX(j, k) Then
            If k > 1 Then
              'all columns to the left of the k column of ith row
              'should be equal or more than corresponding
              'columns of jth row to allow swap
              SWITCH_FLAG = False
              For h = 1 To k - 1
                If DATA_MATRIX(i, h) < DATA_MATRIX(j, h) Then: _
                    SWITCH_FLAG = True
              Next h
              If SWITCH_FLAG = False Then: COMP_FLAG = True
            Else
              'the first column of ith row is more than first
              'col of jth row =>allow swap
              COMP_FLAG = True
            End If
          End If
        Next k
        If COMP_FLAG = True Then
          
          ReDim CTEMP_VECTOR(1 To NCOLUMNS, 1 To 1)
          For l = 1 To NCOLUMNS
            CTEMP_VECTOR(l, 1) = DATA_MATRIX(j, l)
          Next l
          ATEMP_VECTOR = CTEMP_VECTOR
          
          ReDim CTEMP_VECTOR(1 To NCOLUMNS, 1 To 1)
          For l = 1 To NCOLUMNS
            CTEMP_VECTOR(l, 1) = DATA_MATRIX(i, l)
          Next l
          BTEMP_VECTOR = CTEMP_VECTOR
           
           For l = 1 To NCOLUMNS
                DATA_MATRIX(j, l) = BTEMP_VECTOR(l, 1)
           Next l
           For l = 1 To NCOLUMNS
                DATA_MATRIX(i, l) = ATEMP_VECTOR(l, 1)
           Next l
        End If
    Next j
Next i

MATRIX_DOUBLE_SORT_FUNC = DATA_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_DOUBLE_SORT_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_SORT_COLUMNS_FUNC

'DESCRIPTION   : Sorts a given matrix in ascending order and upto a a number os
'columns specified by SCOLUMN

'LIBRARY       : MATRIX
'GROUP         : SORT
'ID            : 012
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_SORT_COLUMNS_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal SCOLUMN As Long = 0)

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim SROW As Long
Dim NROWS As Long

Dim ATEMP_VECTOR As Variant
Dim BTEMP_VECTOR As Variant
Dim DATA_MATRIX As Variant

Dim TEMP_FLAG As Boolean
Dim COMPARE_FLAG As Boolean

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
SROW = LBound(DATA_MATRIX, 1)
NROWS = UBound(DATA_MATRIX, 1)

If SCOLUMN = 0 Then: SCOLUMN = LBound(DATA_MATRIX, 2)

For i = SROW To NROWS - 1
    For j = i + 1 To NROWS
        
        COMPARE_FLAG = False
        
        For k = 1 To SCOLUMN
          If DATA_MATRIX(i, k) > DATA_MATRIX(j, k) Then
            If k > 1 Then
              'all columns to the left of the k column of ith row
              'should be equal or more than corresponding
              'columns of jth row to allow swap
              
              TEMP_FLAG = False
              For l = 1 To k - 1
                If DATA_MATRIX(i, l) < DATA_MATRIX(j, l) Then: TEMP_FLAG = True
              Next l
              If TEMP_FLAG = False Then: COMPARE_FLAG = True
            Else
              'the first column of ith row is more than first
              'col of jth row =>allow swap
              COMPARE_FLAG = True
            End If
          End If
        Next k
        

        If COMPARE_FLAG = True Then
            ReDim ATEMP_VECTOR(1 To UBound(DATA_MATRIX, 2), 1 To 1)
            ReDim BTEMP_VECTOR(1 To UBound(DATA_MATRIX, 2), 1 To 1)
            For h = 1 To UBound(DATA_MATRIX, 2)
              ATEMP_VECTOR(h, 1) = DATA_MATRIX(j, h)
              BTEMP_VECTOR(h, 1) = DATA_MATRIX(i, h)
            Next h
            For h = 1 To UBound(DATA_MATRIX, 2)
              DATA_MATRIX(j, h) = BTEMP_VECTOR(h, 1)
              DATA_MATRIX(i, h) = ATEMP_VECTOR(h, 1)
            Next h
        End If
    Next j
Next i

MATRIX_SORT_COLUMNS_FUNC = DATA_MATRIX
    
Exit Function
ERROR_LABEL:
MATRIX_SORT_COLUMNS_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_QUICK_SORT_FUNC

'DESCRIPTION   : Quick sort of array x which is x(l:r) sorts array x[l..r,n]
'using the mth column for comparison n is the number of columns in the array x.

'LIBRARY       : MATRIX
'GROUP         : SORT
'ID            : 013
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************


Function MATRIX_QUICK_SORT_FUNC(ByRef DATA_RNG As Variant, _
Optional ByRef SCOLUMN As Long = 1, _
Optional ByVal SORT_TYPE As Integer = 1)

' SCOLUMN is the column that will be used for comparison is the column that
' is sorted from low to high)

'If SORT_TYPE is 0 (zero) or omitted, the list is sorted in descending order.
'If SORT_TYPE is any nonzero value, the list is sorted in ascending order.

Dim SROW As Long
Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_FLAG As Boolean
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG

SROW = LBound(DATA_MATRIX, 1)
NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

If SORT_TYPE <> 1 Then: SORT_TYPE = 0

'--------------------------------------------------------------------------------
Select Case SORT_TYPE
'--------------------------------------------------------------------------------
Case 0 '-------------------------DESCENDING
'--------------------------------------------------------------------------------
    TEMP_FLAG = MATRIX_MULT_SORT_FUNC(DATA_MATRIX, SROW, NROWS, SCOLUMN, NCOLUMNS)
    If TEMP_FLAG = False Then: GoTo ERROR_LABEL
    DATA_MATRIX = MATRIX_REVERSE_FUNC(DATA_MATRIX)
'--------------------------------------------------------------------------------
Case Else '----------------------ASCENDING
'--------------------------------------------------------------------------------
    TEMP_FLAG = MATRIX_MULT_SORT_FUNC(DATA_MATRIX, SROW, NROWS, SCOLUMN, NCOLUMNS)
    If TEMP_FLAG = False Then: GoTo ERROR_LABEL
'--------------------------------------------------------------------------------
End Select
'--------------------------------------------------------------------------------

MATRIX_QUICK_SORT_FUNC = DATA_MATRIX
Exit Function
ERROR_LABEL:
MATRIX_QUICK_SORT_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_MULT_SORT_FUNC
'DESCRIPTION   : Support routine for MATRIX_QUICK_SORT_FUNC
'LIBRARY       : MATRIX
'GROUP         : SORT
'ID            : 014
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Private Function MATRIX_MULT_SORT_FUNC(ByRef DATA_MATRIX As Variant, _
ByVal SROW As Long, _
ByVal NROWS As Long, _
ByVal SCOLUMN As Long, _
ByVal NCOLUMNS As Long)
    
Dim ii As Long
Dim jj As Long 'index

Dim PIVOT_VAL As Variant
Dim TEMP_FLAG As Boolean

On Error GoTo ERROR_LABEL

MATRIX_MULT_SORT_FUNC = False
ii = MATRIX_MULT_SORT_PIVOT_FUNC(DATA_MATRIX, SROW, NROWS, SCOLUMN)

If (ii <> -1) Then
    PIVOT_VAL = DATA_MATRIX(ii, SCOLUMN)
    TEMP_FLAG = MATRIX_MULT_SORT_PART_FUNC(DATA_MATRIX, SROW, NROWS, PIVOT_VAL, jj, SCOLUMN, NCOLUMNS)
    If TEMP_FLAG = False Then: GoTo ERROR_LABEL
    TEMP_FLAG = MATRIX_MULT_SORT_FUNC(DATA_MATRIX, SROW, jj - 1, SCOLUMN, NCOLUMNS)
    If TEMP_FLAG = False Then: GoTo ERROR_LABEL
    TEMP_FLAG = MATRIX_MULT_SORT_FUNC(DATA_MATRIX, jj, NROWS, SCOLUMN, NCOLUMNS)
    If TEMP_FLAG = False Then: GoTo ERROR_LABEL
End If

MATRIX_MULT_SORT_FUNC = True

Exit Function
ERROR_LABEL:
MATRIX_MULT_SORT_FUNC = False
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_MULT_SORT_PART_FUNC
'DESCRIPTION   : Support routine for MATRIX_QUICK_SORT_FUNC
'LIBRARY       : MATRIX
'GROUP         : SORT
'ID            : 015
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Private Function MATRIX_MULT_SORT_PART_FUNC(ByRef DATA_MATRIX As Variant, _
ByRef SROW As Long, _
ByRef NROWS As Long, _
ByRef PIVOT_VAL As Variant, _
ByRef jj As Long, _
ByRef SCOLUMN As Long, _
ByRef NCOLUMNS As Long)

Dim ii As Long 'low
Dim kk As Long 'high
Dim ll As Long
Dim TEMP_VAL As Variant

On Error GoTo ERROR_LABEL

MATRIX_MULT_SORT_PART_FUNC = False

ii = SROW
kk = NROWS

Do While ii <= kk
    If DATA_MATRIX(ii, SCOLUMN) < PIVOT_VAL Then
        ii = ii + 1
    ElseIf DATA_MATRIX(kk, SCOLUMN) >= PIVOT_VAL Then
        kk = kk - 1
    Else
        For ll = 1 To NCOLUMNS
            TEMP_VAL = DATA_MATRIX(kk, ll)
            DATA_MATRIX(kk, ll) = DATA_MATRIX(ii, ll)
            DATA_MATRIX(ii, ll) = TEMP_VAL
        Next ll
        kk = kk - 1
        ii = ii + 1
    End If
Loop

jj = ii

MATRIX_MULT_SORT_PART_FUNC = True

Exit Function
ERROR_LABEL:
MATRIX_MULT_SORT_PART_FUNC = False
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_MULT_SORT_PIVOT_FUNC
'DESCRIPTION   : Support routine for MATRIX_QUICK_SORT_FUNC
'LIBRARY       : MATRIX
'GROUP         : SORT
'ID            : 016
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Private Function MATRIX_MULT_SORT_PIVOT_FUNC(ByRef DATA_MATRIX As Variant, _
ByRef SROW As Long, _
ByRef NROWS As Long, _
ByRef SCOLUMN As Long)
    
Dim ii As Long
Dim jj As Long
Dim kk As Long 'index
Dim ll As Long 'mid index
    
Dim iii As Long 'first check
Dim jjj As Long 'second check
    
Dim TEMP_FLAG As Boolean 'is a check if there are at least two
'unequal entries in the list

Dim TEMP1_VAL As Variant
Dim TEMP2_VAL As Variant
Dim TEMP3_VAL As Variant

On Error GoTo ERROR_LABEL

TEMP_FLAG = False
ii = SROW
    
Do While TEMP_FLAG = False And ii < NROWS
    If DATA_MATRIX(ii, SCOLUMN) <> DATA_MATRIX(ii + 1, SCOLUMN) Then
        TEMP_FLAG = True
        jj = ii + 1
    Else
        ii = ii + 1
    End If
Loop
    
kk = ii
If ii < NROWS Then
    If DATA_MATRIX(ii + 1, SCOLUMN) > DATA_MATRIX(ii, SCOLUMN) Then: kk = ii + 1
End If
' kk is needed in case we have three equal pivot values
' in the median of three method

If TEMP_FLAG = True Then
  If NROWS - SROW > 1 Then
     TEMP1_VAL = DATA_MATRIX(SROW, SCOLUMN)
     TEMP2_VAL = DATA_MATRIX(NROWS, SCOLUMN)
     ll = (SROW + NROWS) / 2
     TEMP3_VAL = DATA_MATRIX(ll, SCOLUMN)
' The following subroutine checks to see which element is smaller
     jjj = 0
     iii = 1
     
     If TEMP1_VAL = TEMP2_VAL Then
        If TEMP1_VAL > TEMP3_VAL Then
            jj = SROW  ' pick TEMP1_VAL because it is
                 'larger than TEMP3_VAL
        Else
            jj = ll ' pick TEMP3_VAL because it may be
                 'bigger than TEMP1_VAL and TEMP2_VAL
            If TEMP1_VAL = TEMP3_VAL Then jjj = 1 ' This
             'says that all three elements are equal, so we
             'need to discard the result
        
        End If
     ElseIf TEMP1_VAL = TEMP3_VAL Then
        If TEMP1_VAL > TEMP2_VAL Then
            jj = SROW  ' pick TEMP1_VAL because it is bigger than TEMP2_VAL
        Else
            jj = NROWS  ' pick TEMP2_VAL because we know it is bigger
                 'than TEMP1_VAL and TEMP3_VAL
        End If
     ElseIf TEMP2_VAL = TEMP3_VAL Then
        If TEMP2_VAL > TEMP1_VAL Then
            jj = NROWS  ' pick TEMP2_VAL
        Else
            jj = SROW  ' pick TEMP1_VAL because we know that TEMP1_VAL <> TEMP2_VAL
        End If
     Else
        iii = 0
     End If

     If iii = 0 Then
        If TEMP1_VAL > TEMP2_VAL Then
            If TEMP1_VAL > TEMP3_VAL Then
                If TEMP2_VAL > TEMP3_VAL Then
                    jj = NROWS
                Else
                    jj = ll
                End If
            Else
                jj = SROW
            End If
        ElseIf TEMP2_VAL > TEMP3_VAL Then
            If TEMP1_VAL > TEMP3_VAL Then
                jj = SROW
            Else
                jj = ll
            End If
        Else
            jj = NROWS
        End If
     End If
     ' If iii = 1 we assigned jj already in the
     ' iii for smaller routine
  Else
    jj = kk
    'If DATA_MATRIX(jj, SCOLUMN) < DATA_MATRIX(jj - 1, SCOLUMN) Then
    '   jj = jj - 1
    'End If
  End If
Else
  jj = -1
End If

If jjj = 1 Then: jj = kk

MATRIX_MULT_SORT_PIVOT_FUNC = jj

Exit Function
ERROR_LABEL:
MATRIX_MULT_SORT_PIVOT_FUNC = Err.number
End Function
