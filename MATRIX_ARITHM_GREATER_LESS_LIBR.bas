Attribute VB_Name = "MATRIX_ARITHM_GREATER_LESS_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_ELEMENTS_GREATER_FUNC

'DESCRIPTION   : Returns a binary value indicating whether or not the entries in
'the first array (A) are greater than the second array (B)
'evaluated as an integer.

'LIBRARY       : MATRIX
'GROUP         : ARITHMETIC_GREATER_LESS
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_ELEMENTS_GREATER_FUNC(ByRef DATA1_RNG As Variant, _
ByRef DATA2_RNG As Variant)
  
Dim i As Long
Dim j As Long

Dim SROW As Long
Dim SCOLUMN As Long

Dim NROWS1 As Long
Dim NCOLUMNS1 As Long

Dim NROWS2 As Long
Dim NCOLUMNS2 As Long
  
Dim DATA1_MATRIX As Variant
Dim DATA2_MATRIX As Variant

Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA1_MATRIX = DATA1_RNG
DATA2_MATRIX = DATA2_RNG

NROWS1 = UBound(DATA1_MATRIX, 1)
NROWS2 = UBound(DATA2_MATRIX, 1)

NCOLUMNS1 = UBound(DATA1_MATRIX, 2)
NCOLUMNS2 = UBound(DATA2_MATRIX, 2)

SROW = LBound(DATA1_MATRIX, 1)
SCOLUMN = LBound(DATA1_MATRIX, 2)
 
ReDim TEMP_MATRIX(SROW To NROWS1, SCOLUMN To NCOLUMNS1)
For i = SROW To NROWS1
    For j = SCOLUMN To NCOLUMNS1
        If DATA1_MATRIX(i, j) > DATA2_MATRIX(i, j) Then
            TEMP_MATRIX(i, j) = 1
        Else
            TEMP_MATRIX(i, j) = 0
        End If
    Next j
Next i
MATRIX_ELEMENTS_GREATER_FUNC = TEMP_MATRIX
  
Exit Function
ERROR_LABEL:
MATRIX_ELEMENTS_GREATER_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_ELEMENTS_LESS_FUNC

'DESCRIPTION   : Returns a binary value indicating whether or not the entries in
'the first array (A) are less than the second array (B)
'evaluated as an integer.

'LIBRARY       : MATRIX
'GROUP         : ARITHMETIC_GREATER_LESS
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************


Function MATRIX_ELEMENTS_LESS_FUNC(ByRef DATA1_RNG As Variant, _
ByRef DATA2_RNG As Variant)
  
Dim i As Long
Dim j As Long

Dim SROW As Long
Dim SCOLUMN As Long

Dim NROWS1 As Long
Dim NCOLUMNS1 As Long

Dim NROWS2 As Long
Dim NCOLUMNS2 As Long

Dim DATA1_MATRIX As Variant
Dim DATA2_MATRIX As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA1_MATRIX = DATA1_RNG
DATA2_MATRIX = DATA2_RNG

NROWS1 = UBound(DATA1_MATRIX, 1)
NROWS2 = UBound(DATA2_MATRIX, 1)

NCOLUMNS1 = UBound(DATA1_MATRIX, 2)
NCOLUMNS2 = UBound(DATA2_MATRIX, 2)
 
SROW = LBound(DATA1_MATRIX, 1)
SCOLUMN = LBound(DATA1_MATRIX, 2)
 
ReDim TEMP_MATRIX(SROW To NROWS1, SCOLUMN To NCOLUMNS1)
For i = SROW To NROWS1
    For j = SCOLUMN To NCOLUMNS1
        If DATA1_MATRIX(i, j) < DATA2_MATRIX(i, j) Then
            TEMP_MATRIX(i, j) = 1
        Else
            TEMP_MATRIX(i, j) = 0
        End If
    Next j
Next i
MATRIX_ELEMENTS_LESS_FUNC = TEMP_MATRIX
  
Exit Function
ERROR_LABEL:
MATRIX_ELEMENTS_LESS_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_ELEMENTS_FIND_GREATER_FUNC
'DESCRIPTION   : This function returns "true" when all elements of a given
'ROW are equals to the REF_VALUE
'LIBRARY       : MATRIX
'GROUP         : ARITHMETIC_GREATER_LESS
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009

'************************************************************************************
'************************************************************************************

Function MATRIX_ELEMENTS_FIND_GREATER_FUNC(ByRef DATA_RNG As Variant, _
ByVal SROW As Long, _
ByVal NCOLUMNS As Long, _
Optional ByVal SCOLUMN As Variant = Null, _
Optional ByVal REF_VALUE As Variant = 0, _
Optional ByVal VERSION As Integer = 0)

'----------------------------------DATA_RNG---------------------------------------
'ROW1:  The numeric functions relating to changing the base are set out first.
'ROW2:  The functions that handle text strings come next.
'ROW3:  Two more related functions - SUM_DIGITS and CHECK_DIGITS come last.
'----------------------------------------------------------------------------------

Dim i As Long
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
If VarType(SCOLUMN) = vbNull Then: SCOLUMN = LBound(DATA_MATRIX, 2)

Select Case VERSION
Case 0
    For i = SCOLUMN To NCOLUMNS
        If DATA_MATRIX(SROW, i) > REF_VALUE Then
              MATRIX_ELEMENTS_FIND_GREATER_FUNC = False
              Exit Function
        Else
            MATRIX_ELEMENTS_FIND_GREATER_FUNC = True
        End If
    Next i
Case Else 'This routine checks whether all the digits are
'less than the base.
    MATRIX_ELEMENTS_FIND_GREATER_FUNC = False
    For i = SCOLUMN To NCOLUMNS
        If DATA_MATRIX(SROW, i) > REF_VALUE Then
              MATRIX_ELEMENTS_FIND_GREATER_FUNC = True: Exit Function
        End If
    Next i
End Select
    
Exit Function
ERROR_LABEL:
MATRIX_ELEMENTS_FIND_GREATER_FUNC = Err.number
End Function
