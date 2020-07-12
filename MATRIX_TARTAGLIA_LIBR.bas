Attribute VB_Name = "MATRIX_TARTAGLIA_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_TARTAGLIA_FUNC

'DESCRIPTION   : Returns the (NSIZE x NSIZE) Tartaglia's matrix
'These matrices are hill-conditioned and are useful for testing algorithms

'LIBRARY       : MATRIX
'GROUP         : TARTAGLIA
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_TARTAGLIA_FUNC(ByVal NSIZE As Long)

Dim i As Long
Dim j As Long
Dim k As Long

Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

ReDim TEMP_MATRIX(1 To NSIZE, 1 To NSIZE)

For i = 1 To NSIZE
    For j = 1 To NSIZE
        If i = 1 Then
            TEMP_MATRIX(i, j) = 1
        ElseIf j = 1 Then
            TEMP_MATRIX(i, j) = 1
        ElseIf i < j Then
            For k = 1 To i
                TEMP_MATRIX(i, j) = TEMP_MATRIX(i, j) + TEMP_MATRIX(k, j - 1)
            Next k
        Else
            For k = 1 To j
                TEMP_MATRIX(i, j) = TEMP_MATRIX(i, j) + TEMP_MATRIX(i - 1, k)
            Next k
        End If
    Next j
Next i

MATRIX_TARTAGLIA_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_TARTAGLIA_FUNC = Err.number
End Function
