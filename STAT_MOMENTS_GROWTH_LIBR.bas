Attribute VB_Name = "STAT_MOMENTS_GROWTH_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_AVERAGE_RETURNS_FUNC
'DESCRIPTION   : Compute means of historical data
'LIBRARY       : STATISTICS
'GROUP         : GROWTH
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/21/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_AVERAGE_RETURNS_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal DATA_TYPE As Integer = 0, _
Optional ByVal LOG_SCALE As Integer = 0)

Dim i As Long
Dim j As Long
Dim k As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_SUM As Double
Dim TEMP_VECTOR As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
If DATA_TYPE <> 0 Then DATA_MATRIX = MATRIX_PERCENT_FUNC(DATA_MATRIX, LOG_SCALE)

NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

ReDim TEMP_VECTOR(1 To NCOLUMNS, 1 To 1)

For j = 1 To NCOLUMNS
    TEMP_SUM = 0
    k = 0
    For i = 1 To NROWS
        If IsNumeric(DATA_MATRIX(i, j)) And Not IsEmpty(DATA_MATRIX(i, j)) Then
            TEMP_SUM = TEMP_SUM + DATA_MATRIX(i, j)
            k = k + 1
        End If
    Next i
    If k > 0 Then
        TEMP_VECTOR(j, 1) = TEMP_SUM / k
    Else
        TEMP_VECTOR(j, 1) = "N/A"
    End If
Next j

MATRIX_AVERAGE_RETURNS_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
MATRIX_AVERAGE_RETURNS_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : VECTOR_SHRINK_AVERAGE_RETURNS_FUNC
'DESCRIPTION   :
'LIBRARY       : STATISTICS
'GROUP         : GROWTH
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/21/2009
'************************************************************************************
'************************************************************************************

Function VECTOR_SHRINK_AVERAGE_RETURNS_FUNC(ByRef RETURNS_RNG As Variant, _
ByVal SHRINKAGE_VAL As Double)

Dim i As Long
Dim NSIZE As Long

Dim MEAN_VAL As Double
Dim RETURNS_VECTOR As Variant

On Error GoTo ERROR_LABEL

RETURNS_VECTOR = RETURNS_RNG
If UBound(RETURNS_VECTOR, 1) = 1 Then: _
    RETURNS_VECTOR = MATRIX_TRANSPOSE_FUNC(RETURNS_VECTOR)

NSIZE = UBound(RETURNS_VECTOR, 1)
MEAN_VAL = MATRIX_MEAN_FUNC(RETURNS_VECTOR)(1, 1)
For i = 1 To NSIZE
    RETURNS_VECTOR(i, 1) = RETURNS_VECTOR(i, 1) + SHRINKAGE_VAL * (MEAN_VAL - RETURNS_VECTOR(i, 1))
Next i
VECTOR_SHRINK_AVERAGE_RETURNS_FUNC = RETURNS_VECTOR

Exit Function
ERROR_LABEL:
VECTOR_SHRINK_AVERAGE_RETURNS_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_PERCENT_FUNC
'DESCRIPTION   : RETURNS AN ARRAY WITH PERCENTAGE CHANGE
'LIBRARY       : STATISTICS
'GROUP         : GROWTH
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/21/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_PERCENT_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal LOG_SCALE As Integer = 0)

Dim j As Long
Dim i As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG

NROWS = UBound(DATA_MATRIX, 1) - 1
NCOLUMNS = UBound(DATA_MATRIX, 2)

ReDim TEMP_MATRIX(1 To NROWS, 1 To NCOLUMNS)

For i = 1 To NROWS
    For j = 1 To NCOLUMNS
        If IsNumeric(DATA_MATRIX(i, j)) = False Then
            TEMP_MATRIX(i, j) = DATA_MATRIX(i, j)
        ElseIf (DATA_MATRIX(i, j) = 0) Or (DATA_MATRIX(i + 1, j) = 0) Or _
                (IsEmpty(Trim(DATA_MATRIX(i, j))) = True) Or _
                (IsEmpty(Trim(DATA_MATRIX(i + 1, j))) = True) Then
            TEMP_MATRIX(i, j) = 0
        Else
            If (LOG_SCALE <> 0) Then
                TEMP_MATRIX(i, j) = Log(DATA_MATRIX(i + 1, j) / DATA_MATRIX(i, j))
            Else
                TEMP_MATRIX(i, j) = (DATA_MATRIX(i + 1, j) / DATA_MATRIX(i, j)) - 1
            End If
        End If
    Next j
Next i

MATRIX_PERCENT_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_PERCENT_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_MEAN_FUNC
'DESCRIPTION   : Returns an array with the average values
'LIBRARY       : STATISTICS
'GROUP         : GROWTH
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/21/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_MEAN_FUNC(ByRef DATA_RNG As Variant)

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_SUM As Double
Dim TEMP_VECTOR As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG

NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

ReDim TEMP_VECTOR(1 To 1, 1 To NCOLUMNS)

For j = 1 To NCOLUMNS
   TEMP_SUM = 0
   For i = 1 To NROWS
      TEMP_SUM = TEMP_SUM + DATA_MATRIX(i, j)
   Next i
   TEMP_VECTOR(1, j) = TEMP_SUM / NROWS
Next j
  
MATRIX_MEAN_FUNC = TEMP_VECTOR
  
Exit Function
ERROR_LABEL:
MATRIX_MEAN_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_DEMEAN_FUNC

'DESCRIPTION   : Returns a matrix by subtracting all entries in each column
'with the average value for the column

'LIBRARY       : STATISTICS
'GROUP         : GROWTH
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/21/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_DEMEAN_FUNC(ByRef DATA_RNG As Variant)

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_VECTOR As Variant
Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

ReDim TEMP_MATRIX(1 To NROWS, 1 To NCOLUMNS)

TEMP_VECTOR = MATRIX_MEAN_FUNC(DATA_MATRIX)

For j = 1 To NCOLUMNS
  For i = 1 To NROWS
    TEMP_MATRIX(i, j) = DATA_MATRIX(i, j) - TEMP_VECTOR(1, j)
  Next i
Next j

MATRIX_DEMEAN_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_DEMEAN_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : VECTOR_QUADRATIC_MEAN_FUNC
'DESCRIPTION   : Quadratic Mean Function
'LIBRARY       : STATISTICS
'GROUP         : GROWTH
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/21/2009
'************************************************************************************
'************************************************************************************

Function VECTOR_QUADRATIC_MEAN_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal VERSION As Integer = 1)

Dim i As Long
Dim NROWS As Long
Dim TEMP_VAL As Double
Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then: DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
NROWS = UBound(DATA_VECTOR, 1)

Select Case VERSION
Case 0 'geometric mean of NROWS-numbers
    TEMP_VAL = 1
    For i = 1 To NROWS
        TEMP_VAL = TEMP_VAL * (DATA_VECTOR(i, 1) ^ (1 / NROWS))
    Next i
    VECTOR_QUADRATIC_MEAN_FUNC = TEMP_VAL
Case Else 'quadratic mean of NROWS-numbers
    TEMP_VAL = 0
    For i = 1 To NROWS
        TEMP_VAL = TEMP_VAL + DATA_VECTOR(i, 1) ^ 2
    Next i
    VECTOR_QUADRATIC_MEAN_FUNC = Sqr(TEMP_VAL / NROWS)
End Select
  
Exit Function
ERROR_LABEL:
VECTOR_QUADRATIC_MEAN_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_NORMALIZED_MEAN_FUNC
'DESCRIPTION   : RETURNS AN ARRAY WITH NORMALIZED RETURNS (Z-SCORE)
'LIBRARY       : STATISTICS
'GROUP         : GROWTH
'ID            : 007
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/21/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_NORMALIZED_MEAN_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal DATA_TYPE As Integer = 0, _
Optional ByVal LOG_SCALE As Integer = 0)

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim MEAN_VECTOR As Variant
Dim SIGMA_VECTOR As Variant

Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant
  
On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG

If DATA_TYPE <> 0 Then DATA_MATRIX = MATRIX_PERCENT_FUNC(DATA_MATRIX, LOG_SCALE)

NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

MEAN_VECTOR = MATRIX_MEAN_FUNC(DATA_MATRIX)
SIGMA_VECTOR = MATRIX_STDEVP_FUNC(DATA_MATRIX)

ReDim TEMP_MATRIX(1 To NROWS, 1 To NCOLUMNS)

For j = 1 To NCOLUMNS
   For i = 1 To NROWS
      TEMP_MATRIX(i, j) = (DATA_MATRIX(i, j) - MEAN_VECTOR(1, j)) / SIGMA_VECTOR(1, j)
   Next i
Next j
  
MATRIX_NORMALIZED_MEAN_FUNC = TEMP_MATRIX
    
Exit Function
ERROR_LABEL:
MATRIX_NORMALIZED_MEAN_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_ABSOLUTE_DEVIATION_FUNC
'DESCRIPTION   : Computes mean absolute deviation from median of a data vector
'LIBRARY       : STATISTICS
'GROUP         : GROWTH
'ID            : 008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/21/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_ABSOLUTE_DEVIATION_FUNC(ByRef DATA_RNG As Variant)
    
Dim i As Long
Dim NROWS As Long

Dim TEMP_SUM As Double
Dim TEMP_MEDIAN As Double

Dim DATA_VECTOR As Variant
    
On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then: DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)

TEMP_MEDIAN = HISTOGRAM_PERCENTILE_FUNC(STRIP_NUMERICS_FUNC(DATA_VECTOR), 0.5, 1)
NROWS = UBound(DATA_VECTOR, 1)
TEMP_SUM = 0
For i = 1 To NROWS
    If IsNumeric(DATA_VECTOR(i, 1)) And Not IsEmpty(DATA_VECTOR(i, 1)) Then
        TEMP_SUM = TEMP_SUM + Abs(DATA_VECTOR(i, 1) - TEMP_MEDIAN)
    End If
Next i

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then: DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
MATRIX_ABSOLUTE_DEVIATION_FUNC = TEMP_SUM / VECTOR_COUNT_NUMERICS_FUNC(DATA_VECTOR)

Exit Function
ERROR_LABEL:
MATRIX_ABSOLUTE_DEVIATION_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_MOVING_AVERAGE_FUNC
'DESCRIPTION   : RETURNS A VECTOR WITH MOVING AVERAGES
'LIBRARY       : STATISTICS
'GROUP         : GROWTH
'ID            : 009
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/21/2009
'************************************************************************************
'************************************************************************************


Function MATRIX_MOVING_AVERAGE_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal MA_FACTOR As Long = 3, _
Optional ByVal DATA_TYPE As Integer = 0, _
Optional ByVal LOG_SCALE As Integer = 0)

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP1_SUM As Double
Dim TEMP2_SUM As Double
Dim TEMP3_SUM As Double

Dim TEMP_VECTOR As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
If DATA_TYPE <> 0 Then DATA_MATRIX = MATRIX_PERCENT_FUNC(DATA_MATRIX, LOG_SCALE)
  
NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

If (MA_FACTOR = 1) Or (MA_FACTOR = 0) Or (MA_FACTOR = 2) Then: GoTo ERROR_LABEL
If MA_FACTOR > NROWS Then MA_FACTOR = NROWS

ReDim TEMP_VECTOR(1 To 1, 1 To NCOLUMNS)

For j = 1 To NCOLUMNS
    k = 0
    TEMP1_SUM = 0: TEMP2_SUM = 0: TEMP3_SUM = 0
    For i = 1 To NROWS
      If i <= MA_FACTOR Then
          TEMP3_SUM = TEMP3_SUM + DATA_MATRIX(i, j)
          k = k + 1
          TEMP1_SUM = TEMP1_SUM + TEMP3_SUM / k
      Else
          h = i - MA_FACTOR
          TEMP3_SUM = TEMP3_SUM - DATA_MATRIX(h, j)
          TEMP3_SUM = TEMP3_SUM + DATA_MATRIX(i, j)
          TEMP2_SUM = TEMP2_SUM + TEMP3_SUM / MA_FACTOR
      End If
    Next i
   TEMP_VECTOR(1, j) = (TEMP1_SUM + TEMP2_SUM) / (NROWS)
Next j

MATRIX_MOVING_AVERAGE_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
MATRIX_MOVING_AVERAGE_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : VECTOR_MOVING_AVERAGE_FUNC
'DESCRIPTION   : RETURNS A VECTOR WITH MOVING AVERAGES
'LIBRARY       : STATISTICS
'GROUP         : GROWTH
'ID            : 010
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/21/2009
'************************************************************************************
'************************************************************************************

Function VECTOR_MOVING_AVERAGE_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal MA_FACTOR As Long = 3)

Dim i As Long
Dim j As Long
Dim k As Long
Dim NROWS As Long

Dim TEMP_SUM As Double
Dim DATA_VECTOR As Variant
Dim TEMP_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then
    DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
End If
NROWS = UBound(DATA_VECTOR, 1)

If (MA_FACTOR = 1) Or (MA_FACTOR = 0) Or (MA_FACTOR = 2) Then: GoTo ERROR_LABEL
If MA_FACTOR > NROWS Then MA_FACTOR = NROWS

ReDim TEMP_VECTOR(1 To NROWS, 1 To 1)

j = 0
TEMP_SUM = 0
For i = 1 To NROWS
  If i <= MA_FACTOR Then
      TEMP_SUM = TEMP_SUM + DATA_VECTOR(i, 1)
      j = j + 1
      TEMP_VECTOR(i, 1) = TEMP_SUM / j
  Else
      k = i - MA_FACTOR
      TEMP_SUM = TEMP_SUM - DATA_VECTOR(k, 1)
      TEMP_SUM = TEMP_SUM + DATA_VECTOR(i, 1)
      TEMP_VECTOR(i, 1) = TEMP_SUM / MA_FACTOR
  End If
Next i

VECTOR_MOVING_AVERAGE_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
VECTOR_MOVING_AVERAGE_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_GROWTH_FUNC
'DESCRIPTION   : RETURNS ARRAY WITH GROWTH RATES
'LIBRARY       : STATISTICS
'GROUP         : GROWTH
'ID            : 011
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/21/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_GROWTH_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal LOG_SCALE As Integer = 0, _
Optional ByVal SROW As Long = 1)

Dim j As Long
Dim i As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = (DATA_RNG)

NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

ReDim TEMP_MATRIX(1 To NROWS, 1 To NCOLUMNS)

For j = 1 To NCOLUMNS
    For i = NROWS To 1 Step -1
        If (DATA_MATRIX(i, j) = 0) Then
            TEMP_MATRIX(i, j) = 0
        ElseIf (LOG_SCALE <> 0) Then
            TEMP_MATRIX(i, j) = Log(DATA_MATRIX(i, j) / DATA_MATRIX(SROW, j))
        Else
            TEMP_MATRIX(i, j) = (DATA_MATRIX(i, j) / DATA_MATRIX(SROW, j)) - 1
        End If
    Next i
Next j

MATRIX_GROWTH_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_GROWTH_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_CAGR_FUNC
'DESCRIPTION   : RETURNS A VECTOR WITH COMPOUNDED ANNUAL GROWTH
'LIBRARY       : STATISTICS
'GROUP         : GROWTH
'ID            : 012
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/21/2009
'************************************************************************************
'************************************************************************************


Function MATRIX_CAGR_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal SROW As Long = 1)

Dim j As Long
Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_VECTOR As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG

NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

If (SROW > NROWS) Then: GoTo ERROR_LABEL

ReDim TEMP_VECTOR(1 To 1, 1 To NCOLUMNS)

For j = 1 To NCOLUMNS
    TEMP_VECTOR(1, j) = (DATA_MATRIX(NROWS, j) / DATA_MATRIX(SROW, j)) ^ (1 / (NROWS - SROW)) - 1
Next j

MATRIX_CAGR_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
MATRIX_CAGR_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_GMS_FUNC

'DESCRIPTION   :
'http://www.gummy-stuff.org/magic_sum.htm
'http://www.gummy-stuff.org/distributions-stuff.htm

'LIBRARY       : STATISTICS
'GROUP         : GROWTH
'ID            : 013
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/21/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_GMS_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal DATA_TYPE As Integer = 0)

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_SUM As Double
Dim MULT_VAL As Double
Dim RETURN_VAL As Double

Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
If UBound(DATA_MATRIX, 1) = 1 Then
    DATA_MATRIX = MATRIX_TRANSPOSE_FUNC(DATA_MATRIX)
End If
If DATA_TYPE <> 0 Then: DATA_MATRIX = MATRIX_PERCENT_FUNC(DATA_MATRIX, 0)

NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

ReDim TEMP_MATRIX(1 To 1, 1 To NCOLUMNS)

For j = 1 To NCOLUMNS
    MULT_VAL = 1: TEMP_SUM = 0
    For i = 1 To NROWS
        RETURN_VAL = DATA_MATRIX(i, j)
        MULT_VAL = MULT_VAL * (1 + RETURN_VAL)
        TEMP_SUM = TEMP_SUM + (1 / MULT_VAL)
    Next i
    TEMP_MATRIX(1, j) = TEMP_SUM
Next j
MATRIX_GMS_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_GMS_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : VECTOR_GEOMETRIC_GROWTH_FUNC
'DESCRIPTION   : GEOMETRIC GROWTH
'LIBRARY       : STATISTICS
'GROUP         : GROWTH
'ID            : 014
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/21/2009
'************************************************************************************
'************************************************************************************

Function VECTOR_GEOMETRIC_GROWTH_FUNC(ByVal INIT_VAL As Double, _
ByRef GROWTH_RNG As Variant, _
Optional ByVal VERSION As Integer = 0)

Dim i As Long
Dim NROWS As Long

Dim TEMP_SUM As Double
Dim GROWTH_VAL As Double

Dim DATA_VECTOR As Variant
Dim TEMP_VECTOR As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = GROWTH_RNG
If UBound(DATA_VECTOR, 1) = 1 Then: DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
NROWS = UBound(DATA_VECTOR, 1)

TEMP_SUM = 0
ReDim TEMP_MATRIX(0 To NROWS, 1 To 4)

TEMP_MATRIX(0, 3) = INIT_VAL 'ARITHMETIC GROWTH
TEMP_MATRIX(0, 4) = INIT_VAL 'GEOMETRIC GROWTH

For i = 1 To NROWS
    TEMP_MATRIX(i, 1) = i
    TEMP_MATRIX(i, 2) = DATA_VECTOR(i, 1)
    TEMP_SUM = TEMP_SUM + TEMP_MATRIX(i, 2)
    TEMP_MATRIX(i, 3) = (1 + DATA_VECTOR(i, 1)) * TEMP_MATRIX(i - 1, 3)
    GROWTH_VAL = Exp(Log(TEMP_MATRIX(i, 3) / TEMP_MATRIX(0, 3)) / i) - 1
Next i

For i = 1 To NROWS
    TEMP_MATRIX(i, 4) = (1 + GROWTH_VAL) * TEMP_MATRIX(i - 1, 4)
Next i

'----------------------------------------------------------------------------------
Select Case VERSION
'----------------------------------------------------------------------------------
Case 0
'----------------------------------------------------------------------------------
    ReDim TEMP_VECTOR(1 To 3, 1 To 2)
    
    TEMP_VECTOR(1, 1) = "ARITHMETIC AVERAGE"
    TEMP_VECTOR(1, 2) = TEMP_SUM / NROWS
    
    TEMP_VECTOR(2, 1) = "GEOMETRIC AVERAGE"
    TEMP_VECTOR(2, 2) = GROWTH_VAL
    
    TEMP_VECTOR(3, 1) = "DIFF"
    TEMP_VECTOR(3, 2) = Abs(TEMP_VECTOR(1, 2) - TEMP_VECTOR(2, 2))

    VECTOR_GEOMETRIC_GROWTH_FUNC = TEMP_VECTOR
'----------------------------------------------------------------------------------
Case Else
'----------------------------------------------------------------------------------
    VECTOR_GEOMETRIC_GROWTH_FUNC = TEMP_MATRIX
'----------------------------------------------------------------------------------
End Select
'----------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
VECTOR_GEOMETRIC_GROWTH_FUNC = Err.number
End Function
