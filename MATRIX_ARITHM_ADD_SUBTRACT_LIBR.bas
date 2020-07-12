Attribute VB_Name = "MATRIX_ARITHM_ADD_SUBTRACT_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_ELEMENTS_CUMULATIVE_SUM_FUNC
'DESCRIPTION   : Returns the sum of the entries inside an array
'LIBRARY       : STATISTICS
'GROUP         : ARITHMETIC_ADD_SUBTRACT
'ID            : 00X
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_ELEMENTS_CUMULATIVE_SUM_FUNC(ByRef DATA_RNG As Variant)

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_SUM As Double
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG

NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

TEMP_SUM = 0
For j = 1 To NCOLUMNS
    For i = 1 To NROWS
        TEMP_SUM = TEMP_SUM + DATA_MATRIX(i, j)
    Next i
Next j

MATRIX_ELEMENTS_CUMULATIVE_SUM_FUNC = TEMP_SUM

Exit Function
ERROR_LABEL:
MATRIX_ELEMENTS_CUMULATIVE_SUM_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_ELEMENTS_ADD_SCALAR_FUNC
'DESCRIPTION   : Returns the M = b+A
'LIBRARY       : MATRIX
'GROUP         : ARITHMETIC_ADD_SUBTRACT
'ID            : 00X
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_ELEMENTS_ADD_SCALAR_FUNC(ByRef DATA_RNG As Variant, _
ByVal SCALAR_VAL As Double)
  
  Dim i As Long
  Dim j As Long
  Dim NROWS As Long
  Dim NCOLUMNS As Long
  
  Dim DATA_MATRIX As Variant
  Dim TEMP_MATRIX As Variant
  
  On Error GoTo ERROR_LABEL
  
  DATA_MATRIX = DATA_RNG
  
  NROWS = UBound(DATA_MATRIX, 1)
  NCOLUMNS = UBound(DATA_MATRIX, 2)
  
  ReDim TEMP_MATRIX(1 To NROWS, 1 To NCOLUMNS)
  
  For i = 1 To NROWS
    For j = 1 To NCOLUMNS
          TEMP_MATRIX(i, j) = SCALAR_VAL + DATA_MATRIX(i, j)
    Next j
  Next i
  
  MATRIX_ELEMENTS_ADD_SCALAR_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_ELEMENTS_ADD_SCALAR_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_ELEMENTS_SUBTRACT_SCALAR_FUNC
'DESCRIPTION   : Returns the M = b-A
'LIBRARY       : MATRIX
'GROUP         : ARITHMETIC_ADD_SUBTRACT Subtract & Addition
'ID            : 00X
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_ELEMENTS_SUBTRACT_SCALAR_FUNC(ByRef DATA_RNG As Variant, _
ByVal SCALAR_VAL As Double)
  
  Dim i As Long
  Dim j As Long
  Dim NROWS As Long
  Dim NCOLUMNS As Long
  
  Dim DATA_MATRIX As Variant
  Dim TEMP_MATRIX As Variant
  
  On Error GoTo ERROR_LABEL
  
  DATA_MATRIX = DATA_RNG
  
  NROWS = UBound(DATA_MATRIX, 1)
  NCOLUMNS = UBound(DATA_MATRIX, 2)
  
  ReDim TEMP_MATRIX(1 To NROWS, 1 To NCOLUMNS)
  
  For i = 1 To NROWS
    For j = 1 To NCOLUMNS
          TEMP_MATRIX(i, j) = SCALAR_VAL - DATA_MATRIX(i, j)
    Next j
  Next i
  
  MATRIX_ELEMENTS_SUBTRACT_SCALAR_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_ELEMENTS_SUBTRACT_SCALAR_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_ELEMENTS_ADD_FUNC
'DESCRIPTION   : Returns the M = aA + bB
'LIBRARY       : MATRIX
'GROUP         : ARITHMETIC_ADD_SUBTRACT
'ID            : 00X
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_ELEMENTS_ADD_FUNC(ByRef ADATA_RNG As Variant, _
ByRef BDATA_RNG As Variant, _
Optional ByVal ASCALAR_VAL As Double = 1, _
Optional ByVal BSCALAR_VAL As Double = 1)
  
  Dim i As Long
  Dim j As Long
  
  Dim ANROWS As Long
  Dim ANCOLUMNS As Long
  
  Dim BNROWS As Long
  Dim BNCOLUMNS As Long
  
  Dim ADATA_MATRIX As Variant
  Dim BDATA_MATRIX As Variant
  
  Dim TEMP_MATRIX As Variant

  On Error GoTo ERROR_LABEL

  ADATA_MATRIX = ADATA_RNG
  BDATA_MATRIX = BDATA_RNG

  ANROWS = UBound(ADATA_MATRIX, 1)
  BNROWS = UBound(BDATA_MATRIX, 1)
  
  ANCOLUMNS = UBound(ADATA_MATRIX, 2)
  BNCOLUMNS = UBound(BDATA_MATRIX, 2)
   
'  If (ANROWS <> BNROWS) Or (ANCOLUMNS <> BNCOLUMNS) Then: GoTo ERROR_LABEL

    ReDim TEMP_MATRIX(1 To ANROWS, 1 To ANCOLUMNS)
    For i = 1 To ANROWS
      For j = 1 To ANCOLUMNS
        TEMP_MATRIX(i, j) = (ASCALAR_VAL * ADATA_MATRIX(i, j)) + _
            (BSCALAR_VAL * BDATA_MATRIX(i, j))
      Next j
    Next i
    MATRIX_ELEMENTS_ADD_FUNC = TEMP_MATRIX
  
Exit Function
ERROR_LABEL:
MATRIX_ELEMENTS_ADD_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_ELEMENTS_SUBTRACT_FUNC
'DESCRIPTION   : Returns the M = aA - bB
'LIBRARY       : MATRIX
'GROUP         : ARITHMETIC_ADD_SUBTRACT Subtract & Addition
'ID            : 00X
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_ELEMENTS_SUBTRACT_FUNC(ByRef ADATA_RNG As Variant, _
ByRef BDATA_RNG As Variant, _
Optional ByVal ASCALAR_VAL As Double = 1, _
Optional ByVal BSCALAR_VAL As Double = 1)
  
  Dim i As Long
  Dim j As Long
  
  Dim ANROWS As Long
  Dim ANCOLUMNS As Long
  
  Dim BNROWS As Long
  Dim BNCOLUMNS As Long
  
  Dim ADATA_MATRIX As Variant
  Dim BDATA_MATRIX As Variant
  Dim TEMP_MATRIX As Variant

  On Error GoTo ERROR_LABEL

  ADATA_MATRIX = ADATA_RNG
  BDATA_MATRIX = BDATA_RNG

  ANROWS = UBound(ADATA_MATRIX, 1)
  BNROWS = UBound(BDATA_MATRIX, 1)
  
  ANCOLUMNS = UBound(ADATA_MATRIX, 2)
  BNCOLUMNS = UBound(BDATA_MATRIX, 2)
   
'  If (ANROWS <> BNROWS) Or (ANCOLUMNS <> BNCOLUMNS) Then: GoTo ERROR_LABEL

    ReDim TEMP_MATRIX(1 To ANROWS, 1 To ANCOLUMNS)
    For i = 1 To ANROWS
      For j = 1 To ANCOLUMNS
        TEMP_MATRIX(i, j) = (ASCALAR_VAL * ADATA_MATRIX(i, j)) - _
            (BSCALAR_VAL * BDATA_MATRIX(i, j))
      Next j
    Next i
    MATRIX_ELEMENTS_SUBTRACT_FUNC = TEMP_MATRIX
  
Exit Function
ERROR_LABEL:
MATRIX_ELEMENTS_SUBTRACT_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_ELEMENTS_CONSECUTIVE_SUBTRACT_FUNC

'DESCRIPTION   : Returns a matrix by subtracting all entries in each column
'with the next consecutive entry within that column

'LIBRARY       : MATRIX
'GROUP         : ARITHMETIC_ADD_SUBTRACT
'ID            : 00X
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_ELEMENTS_CONSECUTIVE_SUBTRACT_FUNC(ByRef DATA_RNG As Variant)
 
Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

 DATA_MATRIX = DATA_RNG
 NROWS = UBound(DATA_MATRIX, 1)
 NCOLUMNS = UBound(DATA_MATRIX, 2)

 ReDim TEMP_MATRIX(1 To NROWS - 1, 1 To NCOLUMNS)
 For j = 1 To NCOLUMNS
   For i = 1 To NROWS - 1
     TEMP_MATRIX(i, j) = DATA_MATRIX(i + 1, j) - DATA_MATRIX(i, j)
   Next i
 Next j
 MATRIX_ELEMENTS_CONSECUTIVE_SUBTRACT_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_ELEMENTS_CONSECUTIVE_SUBTRACT_FUNC = Err.number
End Function
