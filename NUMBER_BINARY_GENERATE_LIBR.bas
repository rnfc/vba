Attribute VB_Name = "NUMBER_BINARY_GENERATE_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : BINARY_TABLE_FUNC
'DESCRIPTION   : Generate table of binary numbers
'LIBRARY       : NUMBER_BINARY
'GROUP         : GENERATE
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function BINARY_TABLE_FUNC(ByVal NSIZE As Long, _
Optional ByVal VERSION As Integer = 0, _
Optional ByVal SWITCH_FLAG As Boolean = False)

'NSIZE: Number of Binary variables

'VERSION: If you like, you can indicate if
'the table will be the least significant digit column
'to the right or to the left.

'SWITCH_FLAG: You can choose also the logic symbol “1/0”
'or “TRUE/FALSE”

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim ATEMP_MATRIX As Variant
Dim BTEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

If NSIZE < 1 Then: GoTo ERROR_LABEL

BTEMP_MATRIX = BINARY_GENERATOR_FUNC(NSIZE, VERSION)

NROWS = UBound(BTEMP_MATRIX, 1)
NCOLUMNS = UBound(BTEMP_MATRIX, 2)

'--------------------------------------------------------------------------
If SWITCH_FLAG Then 'change 1-0 with true/false
'--------------------------------------------------------------------------
    ReDim ATEMP_MATRIX(1 To NROWS, 1 To NCOLUMNS)
    For i = 1 To NROWS
        For j = 1 To NCOLUMNS
            If BTEMP_MATRIX(i, j) = 0 Then ATEMP_MATRIX(i, j) = False
            If BTEMP_MATRIX(i, j) = 1 Then ATEMP_MATRIX(i, j) = True
        Next j
    Next i
    BINARY_TABLE_FUNC = ATEMP_MATRIX
'--------------------------------------------------------------------------
Else
'--------------------------------------------------------------------------
    BINARY_TABLE_FUNC = BTEMP_MATRIX
'--------------------------------------------------------------------------
End If
'--------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
BINARY_TABLE_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : BINARY_GENERATOR_FUNC
'DESCRIPTION   : Binary Generator
'LIBRARY       : NUMBER_BINARY
'GROUP         : GENERATE
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************


Function BINARY_GENERATOR_FUNC(ByVal NSIZE As Long, _
Optional ByVal VERSION As Integer = 0)

'NSIZE: Number of Binary variables

'VERSION: If you like, you can indicate if
'the table will be the least significant digit column
'to the right or to the left.

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim TEMP_VALUE As Variant
Dim ATEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

i = 2 ^ NSIZE
ReDim ATEMP_MATRIX(1 To i, 1 To NSIZE)

'-----------------------------------------------------------------------
For k = 2 To i
'-----------------------------------------------------------------------
    h = 1
    ATEMP_MATRIX(k, 1) = 1
'-----------------------------------------------------------------------
    For j = 1 To NSIZE
'-----------------------------------------------------------------------
        If ATEMP_MATRIX(k - 1, j) = 1 Then
            If h = 1 Then
                ATEMP_MATRIX(k, j) = 0
            Else
                ATEMP_MATRIX(k, j) = 1
            End If
'-----------------------------------------------------------------------
        Else
'-----------------------------------------------------------------------
            If h = 1 Then
                ATEMP_MATRIX(k, j) = 1
                h = 0
            End If
        End If
'-----------------------------------------------------------------------
    Next j
'-----------------------------------------------------------------------
'-----------------------------------------------------------------------
Next k
'-----------------------------------------------------------------------

'-----------------------------------------------------------------------
If VERSION = 1 Then  'least sign. bit to the right
'-----------------------------------------------------------------------
    For k = 1 To i
        For j = 1 To Int(NSIZE / 2)
            l = NSIZE + 1 - j
            TEMP_VALUE = ATEMP_MATRIX(k, j)
            ATEMP_MATRIX(k, j) = ATEMP_MATRIX(k, l)
            ATEMP_MATRIX(k, l) = TEMP_VALUE
        Next j
    Next k
'-----------------------------------------------------------------------
End If
'-----------------------------------------------------------------------

BINARY_GENERATOR_FUNC = ATEMP_MATRIX

Exit Function
ERROR_LABEL:
BINARY_GENERATOR_FUNC = Err.number
End Function


