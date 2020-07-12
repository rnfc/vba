Attribute VB_Name = "MATRIX_ARITHM_MIN_MAX_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
                    
'************************************************************************************
'************************************************************************************
'FUNCTION      : MAX_FUNC
'DESCRIPTION   : RETURNS THE MAX VALUE WITHIN AN ARRAY
'LIBRARY       : MATRIX
'GROUP         : ARITHMETIC_MIN_MAX
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_ELEMENTS_MAX_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal VERSION As Integer = 0)

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_VAL As Variant
Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
 
NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)
    
ReDim TEMP_MATRIX(1 To 1, 1 To NCOLUMNS)

For j = 1 To NCOLUMNS
   TEMP_VAL = DATA_MATRIX(1, j)
   For i = 1 To NROWS
      If TEMP_VAL < DATA_MATRIX(i, j) Then TEMP_VAL = DATA_MATRIX(i, j)
   Next i
   TEMP_MATRIX(1, j) = TEMP_VAL
Next j
   
TEMP_VAL = TEMP_MATRIX(1, 1)
For i = 1 To NCOLUMNS
   If TEMP_VAL < TEMP_MATRIX(1, i) Then TEMP_VAL = TEMP_MATRIX(1, i)
Next i
  
Select Case VERSION
  Case 0
      MATRIX_ELEMENTS_MAX_FUNC = TEMP_VAL
  Case Else
      MATRIX_ELEMENTS_MAX_FUNC = TEMP_MATRIX
End Select

Exit Function
ERROR_LABEL:
MATRIX_ELEMENTS_MAX_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_ELEMENTS_MIN_FUNC
'DESCRIPTION   : RETURNS THE MIN VALUE WITHIN AN ARRAY
'LIBRARY       : MATRIX
'GROUP         : ARITHMETIC_MIN_MAX
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_ELEMENTS_MIN_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal VERSION As Integer = 0)

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_VAL As Variant
Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG

NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)
    
ReDim TEMP_MATRIX(1 To 1, 1 To NCOLUMNS)

For j = 1 To NCOLUMNS
   TEMP_VAL = DATA_MATRIX(1, j)
   For i = 1 To NROWS
      If TEMP_VAL > DATA_MATRIX(i, j) Then TEMP_VAL = DATA_MATRIX(i, j)
   Next i
   TEMP_MATRIX(1, j) = TEMP_VAL
Next j

TEMP_VAL = TEMP_MATRIX(1, 1)
For i = 1 To NCOLUMNS
   If TEMP_VAL > TEMP_MATRIX(1, i) Then TEMP_VAL = TEMP_MATRIX(1, i)
Next i

Select Case VERSION
  Case 0
    MATRIX_ELEMENTS_MIN_FUNC = TEMP_VAL
  Case Else
    MATRIX_ELEMENTS_MIN_FUNC = TEMP_MATRIX
End Select
  
Exit Function
ERROR_LABEL:
MATRIX_ELEMENTS_MIN_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : VECTOR_ELEMENTS_MIN_FUNC
'DESCRIPTION   : Returns the minimum value in a vector
'LIBRARY       : MATRIX
'GROUP         : MIN_MAX
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function VECTOR_ELEMENTS_MIN_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal tolerance As Double = 2 ^ 52)

Dim i As Long

Dim SROW As Long
Dim NROWS As Long

Dim TEMP_VAL As Double
Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then
    DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
End If

SROW = LBound(DATA_VECTOR, 1)
NROWS = UBound(DATA_VECTOR, 1)

For i = SROW To NROWS
    If DATA_VECTOR(i, 1) < tolerance Then: TEMP_VAL = DATA_VECTOR(i, 1)
Next i

VECTOR_ELEMENTS_MIN_FUNC = TEMP_VAL

Exit Function
ERROR_LABEL:
VECTOR_ELEMENTS_MIN_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : VECTOR_ELEMENTS_MAX_FUNC
'DESCRIPTION   : Returns the maximum value in a vector
'LIBRARY       : MATRIX
'GROUP         : MIN_MAX
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function VECTOR_ELEMENTS_MAX_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal tolerance As Double = -2 ^ 52)

Dim i As Long

Dim SROW As Long
Dim NROWS As Long

Dim TEMP_VAL As Double
Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then
    DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
End If

SROW = LBound(DATA_VECTOR, 1)
NROWS = UBound(DATA_VECTOR, 1)

For i = SROW To NROWS
    If DATA_VECTOR(i, 1) > tolerance Then: TEMP_VAL = DATA_VECTOR(i, 1)
Next i
VECTOR_ELEMENTS_MAX_FUNC = TEMP_VAL

Exit Function
ERROR_LABEL:
VECTOR_ELEMENTS_MAX_FUNC = Err.number
End Function


