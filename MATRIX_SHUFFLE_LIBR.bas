Attribute VB_Name = "MATRIX_SHUFFLE_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : ARRAY_SHUFFLE_FUNC
'DESCRIPTION   :
'In various applications, you may find in useful or necessary to
'randomize an array. That is, to reorder the elements in random
'order. This page describes to VBA procedures to do this. The first
'procedure, ARRAY_SHUFFLE_FUNC, takes an input array and returns a new
'array containing the elements of the input array in random order.
'The contents and order of the input array are not modified. The
'second procedure, ARRAY_PLACE_SHUFFLE_FUNC, randomizes an input array,
'modifying the contents and order of the input array. This procedure
'does not return a value.

'This function returns the values of DATA_ARR in random order. The original
'DATA_ARR is not modified.

'LIBRARY       : MATRIX
'GROUP         : SHUFFLE
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function ARRAY_SHUFFLE_FUNC(ByRef DATA_ARR() As Variant)

Dim i As Long
Dim j As Long

Dim SROW As Long
Dim NROWS As Long
Dim NSIZE As Long

Dim TEMP_VAL As Variant
Dim TEMP_ARR() As Variant

On Error GoTo ERROR_LABEL

SROW = LBound(DATA_ARR)
NROWS = UBound(DATA_ARR)

Randomize
NSIZE = NROWS - SROW + 1
ReDim TEMP_ARR(SROW To NROWS)
For i = SROW To NROWS
    TEMP_ARR(i) = DATA_ARR(i)
Next i
For i = SROW To NROWS
    j = Int((NROWS - SROW + 1) * Rnd + SROW)
    If i <> j Then
        TEMP_VAL = TEMP_ARR(i)
        TEMP_ARR(i) = TEMP_ARR(j)
        TEMP_ARR(j) = TEMP_VAL
    End If
Next i

ARRAY_SHUFFLE_FUNC = TEMP_ARR

Exit Function
ERROR_LABEL:
ARRAY_SHUFFLE_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : ARRAY_PLACE_SHUFFLE_FUNC
'DESCRIPTION   : This shuffles DATA_ARR to random order, randomized in place.
'LIBRARY       : MATRIX
'GROUP         : SHUFFLE
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function ARRAY_PLACE_SHUFFLE_FUNC(ByRef DATA_ARR() As Variant)
    
Dim i As Long
Dim j As Long

Dim SROW As Long
Dim NROWS As Long
Dim NSIZE As Long

Dim TEMP_VAL As Variant

On Error GoTo ERROR_LABEL

ARRAY_PLACE_SHUFFLE_FUNC = False

SROW = LBound(DATA_ARR)
NROWS = UBound(DATA_ARR)

Randomize
NSIZE = NROWS - SROW + 1
For i = SROW To NROWS
    j = Int((NROWS - SROW + 1) * Rnd + SROW)
    If i <> j Then
        TEMP_VAL = DATA_ARR(i)
        DATA_ARR(i) = DATA_ARR(j)
        DATA_ARR(j) = TEMP_VAL
    End If
Next i

ARRAY_PLACE_SHUFFLE_FUNC = True
    
Exit Function
ERROR_LABEL:
ARRAY_PLACE_SHUFFLE_FUNC = False
End Function
