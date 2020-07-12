Attribute VB_Name = "MATRIX_ARITHM_SUM_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_SUM_FUNC
'DESCRIPTION   : adds all the numbers in a matrix/vector
'LIBRARY       : MATRIX_ARITHM
'GROUP         : SUM
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_SUM_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal VERSION As Integer = 0)

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_SUM As Double
Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG

NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

ReDim TEMP_MATRIX(1 To 1, 1 To NCOLUMNS)

'-------------------------------------------------------------------------
Select Case VERSION
'-------------------------------------------------------------------------
Case 0
'-------------------------------------------------------------------------
    For j = 1 To NCOLUMNS
    TEMP_SUM = 0
        For i = 1 To NROWS
            TEMP_SUM = TEMP_SUM + DATA_MATRIX(i, j)
        Next i
        TEMP_MATRIX(1, j) = TEMP_SUM
    Next j
'-------------------------------------------------------------------------
Case Else
'-------------------------------------------------------------------------
    For j = 1 To NCOLUMNS
    TEMP_SUM = 0
        For i = 1 To NROWS
            TEMP_SUM = TEMP_SUM + Abs(DATA_MATRIX(i, j))
        Next i
        TEMP_MATRIX(1, j) = TEMP_SUM
    Next j
'-------------------------------------------------------------------------
End Select
'-------------------------------------------------------------------------

MATRIX_SUM_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_SUM_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_SUM_PRODUCT_FUNC
'DESCRIPTION   : Multiplies corresponding components in the given two arrays, and
'returns the sum of those products.
'LIBRARY       : MATRIX_ARITHM
'GROUP         : SUM
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_SUM_PRODUCT_FUNC(ByRef DATA1_RNG As Variant, _
ByRef DATA2_RNG As Variant)

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_SUM As Double

Dim DATA1_MATRIX As Variant
Dim DATA2_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA1_MATRIX = DATA1_RNG
DATA2_MATRIX = DATA2_RNG
If UBound(DATA1_MATRIX, 1) <> UBound(DATA2_MATRIX, 1) Then: GoTo ERROR_LABEL
If UBound(DATA1_MATRIX, 2) <> UBound(DATA2_MATRIX, 2) Then: GoTo ERROR_LABEL

NROWS = UBound(DATA1_MATRIX, 1)
NCOLUMNS = UBound(DATA1_MATRIX, 2)

TEMP_SUM = 0
For j = 1 To NCOLUMNS
    For i = 1 To NROWS
        TEMP_SUM = TEMP_SUM + DATA1_MATRIX(i, j) * DATA2_MATRIX(i, j)
    Next i
Next j

MATRIX_SUM_PRODUCT_FUNC = TEMP_SUM

Exit Function
ERROR_LABEL:
MATRIX_SUM_PRODUCT_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_MULT_SUM_PRODUCT_FUNC
'DESCRIPTION   : Multiplies corresponding components in the given arrays, and
'returns the sum of those products.
'LIBRARY       : MATRIX_ARITHM
'GROUP         : SUM
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_MULT_SUM_PRODUCT_FUNC(ByRef DATA1_RNG As Variant, _
ByRef DATA2_RNG As Variant, _
ByRef DATA3_RNG As Variant, _
Optional ByRef DATA4_RNG As Variant)

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_SUM As Double

Dim DATA1_MATRIX As Variant
Dim DATA2_MATRIX As Variant
Dim DATA3_MATRIX As Variant
Dim DATA4_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA1_MATRIX = DATA1_RNG
DATA2_MATRIX = DATA2_RNG
DATA3_MATRIX = DATA3_RNG
DATA4_MATRIX = DATA4_RNG

NROWS = UBound(DATA1_MATRIX, 1)
NCOLUMNS = UBound(DATA1_MATRIX, 2)

'--------------------------------------------------------------------------
If IsArray(DATA4_MATRIX) = False Then
'--------------------------------------------------------------------------
    TEMP_SUM = 0
    For j = 1 To NCOLUMNS
        For i = 1 To NROWS
            TEMP_SUM = TEMP_SUM + DATA1_MATRIX(i, j) * DATA2_MATRIX(i, j) * DATA3_MATRIX(i, j)
        Next i
    Next j
'--------------------------------------------------------------------------
Else 'If IsArray(DATA4_MATRIX) = True Then
'--------------------------------------------------------------------------
    TEMP_SUM = 0
    For j = 1 To NCOLUMNS
        For i = 1 To NROWS
            TEMP_SUM = TEMP_SUM + DATA1_MATRIX(i, j) * DATA2_MATRIX(i, j) * DATA3_MATRIX(i, j) * DATA4_MATRIX(i, j)
        Next i
    Next j
'--------------------------------------------------------------------------
End If
'--------------------------------------------------------------------------
MATRIX_MULT_SUM_PRODUCT_FUNC = TEMP_SUM
'--------------------------------------------------------------------------
Exit Function
ERROR_LABEL:
MATRIX_MULT_SUM_PRODUCT_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : RNG_SUMPRODUCT_FUNC
'DESCRIPTION   : Multiplies corresponding components in the given arrays, and
'returns the sum of those products.
'LIBRARY       : MATRIX_ARITHM
'GROUP         : SUM
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function RNG_SUMPRODUCT_FUNC(ParamArray DATA_RNG() As Variant)

Dim i As Long
Dim j As Long

Dim TEMP_SUM As Double
Dim PRODUCT_VAL As Double

On Error GoTo ERROR_LABEL

TEMP_SUM = 0
For i = 1 To DATA_RNG(LBound(DATA_RNG)).Cells.COUNT
    PRODUCT_VAL = 1
    For j = LBound(DATA_RNG) To UBound(DATA_RNG)
        PRODUCT_VAL = PRODUCT_VAL * DATA_RNG(j)(i)
    Next j
    TEMP_SUM = TEMP_SUM + PRODUCT_VAL
Next i

RNG_SUMPRODUCT_FUNC = TEMP_SUM

Exit Function
ERROR_LABEL:
RNG_SUMPRODUCT_FUNC = Err.number
End Function
