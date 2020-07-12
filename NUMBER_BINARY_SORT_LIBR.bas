Attribute VB_Name = "NUMBER_BINARY_SORT_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : BINARY_SORT_FUNC
'DESCRIPTION   : Sort Routine with Swapping Algorithm
'LIBRARY       : NUMBER_BINARY
'GROUP         : SORT
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

'// PERFECT

Function BINARY_SORT_FUNC(ByRef DATA_RNG As Variant, _
Optional ByRef ACOLUMN As Long = 1, _
Optional ByRef VERSION As Integer = 1)

Dim i As Long
Dim j As Long
Dim k As Long

Dim ii As Long
Dim jj As Long
Dim kk As Long

Dim SROW As Long
Dim NROWS As Long

Dim SCOLUMN As Long
Dim NCOLUMNS As Long

Dim DATA_MATRIX As Variant
Dim EXCHANGE_FLAG As Boolean

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG

If VERSION <> 1 Then: VERSION = 0 ' Descending

SROW = LBound(DATA_MATRIX, 1)
NROWS = UBound(DATA_MATRIX, 1)
SCOLUMN = LBound(DATA_MATRIX, 2)
NCOLUMNS = UBound(DATA_MATRIX, 2)
If ACOLUMN = 0 Then ACOLUMN = SCOLUMN

Do
    EXCHANGE_FLAG = False
    For i = SROW To NROWS Step 2
        k = i + 1
        If k > NROWS Then Exit For
        ii = DATA_MATRIX(i, ACOLUMN)
        jj = DATA_MATRIX(k, ACOLUMN)
        If (ii > jj And VERSION = 1) Or _
           (ii < jj And VERSION = 0) Then
            'swap rows
            For j = SCOLUMN To NCOLUMNS
                kk = DATA_MATRIX(k, j)
                DATA_MATRIX(k, j) = DATA_MATRIX(i, j)
                DATA_MATRIX(i, j) = kk
            Next j
            EXCHANGE_FLAG = True
        End If
    Next i
    If SROW = LBound(DATA_MATRIX, SCOLUMN) Then
          SROW = LBound(DATA_MATRIX, SCOLUMN) + 1
    Else
          SROW = LBound(DATA_MATRIX, SCOLUMN)
    End If
Loop Until EXCHANGE_FLAG = False And SROW = LBound(DATA_MATRIX, SCOLUMN)

BINARY_SORT_FUNC = DATA_MATRIX

Exit Function
ERROR_LABEL:
BINARY_SORT_FUNC = Err.number
End Function
