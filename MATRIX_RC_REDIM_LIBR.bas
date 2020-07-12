Attribute VB_Name = "MATRIX_RC_REDIM_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : ARRAY_DIMENSION_FUNC
'DESCRIPTION   : No. Dimensions in array
'LIBRARY       : MATRIX
'GROUP         : RC_REDIM
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function ARRAY_DIMENSION_FUNC(ByRef DATA_RNG As Variant)
Dim i As Long
Dim j As Long
Dim DATA_VECTOR As Variant
On Error Resume Next
DATA_VECTOR = DATA_RNG
i = 0
Do
    i = i + 1
    j = UBound(DATA_VECTOR, i)
Loop Until Err.number <> 0
ARRAY_DIMENSION_FUNC = i - 1
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_REDIM_FUNC
'DESCRIPTION   : Extend the lenght of a 2D array preserving its values
'                k = amount to extend
'LIBRARY       : MATRIX
'GROUP         : RC_REDIM
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_REDIM_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal k As Long = 1, _
Optional ByVal VERSION As Integer = 0)

'VERSION = 0; to Rows
'VERSION = 1; to Columns

Dim i As Long
Dim j As Long

Dim NSIZE As Long

Dim SROW As Long
Dim SCOLUMN As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim DATA1_MATRIX As Variant
Dim DATA2_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA1_MATRIX = DATA_RNG
DATA2_MATRIX = DATA1_MATRIX  'save values in DATA_MATRIX temp array

SROW = LBound(DATA2_MATRIX, 1)
NROWS = UBound(DATA2_MATRIX, 1)

SCOLUMN = LBound(DATA2_MATRIX, 2)
NCOLUMNS = UBound(DATA2_MATRIX, 2)

'-------------------------------------------------------------------------------
Select Case VERSION
'-------------------------------------------------------------------------------
Case 0 ' Redim Preserve Rows
'-------------------------------------------------------------------------------
    NSIZE = NROWS + k
    ReDim DATA1_MATRIX(SROW To NSIZE, SCOLUMN To NCOLUMNS) 'reload values
    For i = SROW To NROWS
        For j = SCOLUMN To NCOLUMNS
            DATA1_MATRIX(i, j) = DATA2_MATRIX(i, j)
        Next j
    Next i
'-------------------------------------------------------------------------------
Case Else ' Redim Preserve Columns
'-------------------------------------------------------------------------------
    NSIZE = NCOLUMNS + k
    ReDim DATA1_MATRIX(SROW To NROWS, SCOLUMN To NSIZE) 'reload values
    For i = SROW To NROWS
        For j = SCOLUMN To NCOLUMNS
            DATA1_MATRIX(i, j) = DATA2_MATRIX(i, j)
        Next j
    Next i
'-------------------------------------------------------------------------------
End Select
'-------------------------------------------------------------------------------

MATRIX_REDIM_FUNC = DATA1_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_REDIM_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_RESIZE_FUNC
'DESCRIPTION   : Resize an array preserving or cutting its content
'                k = amount to extend
'LIBRARY       : MATRIX
'GROUP         : RC_REDIM
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_RESIZE_FUNC(ByRef DATA_RNG As Variant, _
ByVal AROW As Long, _
Optional ByVal ACOLUMN As Long)

Dim i As Long
Dim j As Long

Dim ii As Long 'size rows
Dim jj As Long 'size columns

Dim SROW As Long
Dim SCOLUMN As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG

'------------------------------------------------------------------------
If IsMissing(ACOLUMN) = True Then  'is vector
'------------------------------------------------------------------------
    SROW = LBound(DATA_MATRIX, 1)
    NROWS = UBound(DATA_MATRIX, 1)
    If (NROWS - SROW + 1) = 1 Then
        DATA_MATRIX = MATRIX_TRANSPOSE_FUNC(DATA_MATRIX)
    End If
    TEMP_MATRIX = DATA_MATRIX
    ReDim DATA_MATRIX(SROW To AROW)
    If AROW < (NROWS - SROW + 1) Then
        ii = AROW
    Else
        ii = (NROWS - SROW + 1)
    End If
    If IS_1D_ARRAY_FUNC(TEMP_MATRIX) = True Then
        For i = SROW To ii
            DATA_MATRIX(i) = TEMP_MATRIX(i)
        Next i
    ElseIf IS_2D_ARRAY_FUNC(TEMP_MATRIX) = True Then
        For i = SROW To ii
            DATA_MATRIX(i) = TEMP_MATRIX(i, 1)
        Next i
    Else
        GoTo ERROR_LABEL
    End If
'------------------------------------------------------------------------
Else 'is an array
'------------------------------------------------------------------------
    SROW = LBound(DATA_MATRIX, 1)
    SCOLUMN = LBound(DATA_MATRIX, 2)
    
    NROWS = UBound(DATA_MATRIX, 1)
    NCOLUMNS = UBound(DATA_MATRIX, 2)
    
    TEMP_MATRIX = DATA_MATRIX
        
    ReDim DATA_MATRIX(SROW To AROW, SCOLUMN To ACOLUMN)

    If AROW < (NROWS - SROW + 1) Then
        ii = AROW
    Else
        ii = NROWS
    End If
    
    If ACOLUMN < (NCOLUMNS - SCOLUMN + 1) Then
        jj = ACOLUMN
    Else
        jj = (NCOLUMNS - SCOLUMN + 1)
    End If
    
    For i = SROW To ii
        For j = SCOLUMN To jj
            DATA_MATRIX(i, j) = TEMP_MATRIX(i, j)
        Next j
    Next i
'------------------------------------------------------------------------
End If
'------------------------------------------------------------------------

MATRIX_RESIZE_FUNC = DATA_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_RESIZE_FUNC = Err.number
End Function
