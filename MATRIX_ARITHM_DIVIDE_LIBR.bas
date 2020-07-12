Attribute VB_Name = "MATRIX_ARITHM_DIVIDE_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_ELEMENTS_DIVIDE_FUNC
'DESCRIPTION   : Returns the M = aA / bB
'LIBRARY       : MATRIX
'GROUP         : ARITHMETIC_DIVIDE
'ID            : 00X
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************


Function MATRIX_ELEMENTS_DIVIDE_FUNC(ByRef DATA1_RNG As Variant, _
ByRef DATA2_RNG As Variant, _
Optional ByVal SCALAR1_VAL As Double = 1, _
Optional ByVal SCALAR2_VAL As Double = 1)
  
Dim i As Long
Dim j As Long

Dim NROWS1 As Long
Dim NCOLUMNS1 As Long

Dim NROWS2 As Long
Dim NCOLUMNS2 As Long

Dim TEMP_MATRIX As Variant
Dim DATA1_MATRIX As Variant
Dim DATA2_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA1_MATRIX = DATA1_RNG
DATA2_MATRIX = DATA2_RNG

NROWS1 = UBound(DATA1_MATRIX, 1)
NROWS2 = UBound(DATA2_MATRIX, 1)

NCOLUMNS1 = UBound(DATA1_MATRIX, 2)
NCOLUMNS2 = UBound(DATA2_MATRIX, 2)

'  If (NROWS1 <> NROWS2) Or (NCOLUMNS1 <> NCOLUMNS2) Then: GoTo ERROR_LABEL

ReDim TEMP_MATRIX(1 To NROWS1, 1 To NCOLUMNS1)

For i = 1 To NROWS1
    For j = 1 To NCOLUMNS1
        TEMP_MATRIX(i, j) = (SCALAR1_VAL * DATA1_MATRIX(i, j)) / (SCALAR2_VAL * DATA2_MATRIX(i, j))
    Next j
Next i

MATRIX_ELEMENTS_DIVIDE_FUNC = TEMP_MATRIX
  
Exit Function
ERROR_LABEL:
MATRIX_ELEMENTS_DIVIDE_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_ELEMENTS_DIVIDE_SCALAR_FUNC
'DESCRIPTION   :
'LIBRARY       : MATRIX
'GROUP         : ARITHMETIC_DIVIDE
'ID            : 00X
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_ELEMENTS_DIVIDE_SCALAR_FUNC(ByVal DATA_RNG As Variant, _
ByVal X_VAL As Double)
  
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

ReDim TEMP_MATRIX(1 To NROWS, 1 To NCOLUMNS)

For i = 1 To NROWS
    For j = 1 To NCOLUMNS
        TEMP_MATRIX(i, j) = DATA_MATRIX(i, j) / X_VAL
    Next j
Next i
MATRIX_ELEMENTS_DIVIDE_SCALAR_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_ELEMENTS_DIVIDE_SCALAR_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_ELEMENTS_FRACTION_FUNC
'DESCRIPTION   : Returns the fraction of each entry in a vector (e.g., 1 / x)
'LIBRARY       : MATRIX
'GROUP         : ARITHMETIC_DIVIDE
'ID            : 00X
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************


Function MATRIX_ELEMENTS_FRACTION_FUNC(ByRef DATA_RNG As Variant)

Dim i As Long
Dim NROWS As Long
'Dim TEMP_VECTOR As Variant
Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then: _
DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
NROWS = UBound(DATA_VECTOR, 1)

'ReDim TEMP_VECTOR(1 To NROWS, 1 To 1)
'For i = 1 To NROWS: TEMP_VECTOR(i, 1) = 1 / DATA_VECTOR(i, 1): Next i
For i = 1 To NROWS: DATA_VECTOR(i, 1) = 1 / DATA_VECTOR(i, 1): Next i

MATRIX_ELEMENTS_FRACTION_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
MATRIX_ELEMENTS_FRACTION_FUNC = Err.number
End Function
