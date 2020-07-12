Attribute VB_Name = "MATRIX_CHECK_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_CROSS_CHECK_VECTOR_FUNC

'DESCRIPTION   : Cross check the entries in the reference vector with the entries in
'the source matrix, return the row position of the entries

'LIBRARY       : MATRIX
'GROUP         : CHECK
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_CROSS_CHECK_VECTOR_FUNC(ByRef SRC_MATRIX_RNG As Variant, _
ByRef REFER_VECTOR_RNG As Variant, _
Optional ByVal epsilon As Double = 0)

Dim i As Long
Dim j As Long

Dim SROW As Long
Dim NROWS As Long

Dim SCOLUMN As Long
Dim NCOLUMNS As Long

Dim TEMP_SUM As Double

Dim DATA_VECTOR As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = SRC_MATRIX_RNG
SROW = LBound(DATA_MATRIX, 1)
NROWS = UBound(DATA_MATRIX, 1)
SCOLUMN = LBound(DATA_MATRIX, 2)
NCOLUMNS = UBound(DATA_MATRIX, 2)
DATA_VECTOR = REFER_VECTOR_RNG
If IS_1D_ARRAY_FUNC(DATA_VECTOR) = True Then
    MATRIX_CROSS_CHECK_VECTOR_FUNC = 0
    For i = SROW To NROWS
        TEMP_SUM = 0
        For j = SCOLUMN To NCOLUMNS
            TEMP_SUM = TEMP_SUM + Abs(DATA_VECTOR(j) - DATA_MATRIX(i, j))
        Next j
        If TEMP_SUM <= epsilon Then MATRIX_CROSS_CHECK_VECTOR_FUNC = i: Exit For
    Next i
ElseIf IS_2D_ARRAY_FUNC(DATA_VECTOR) = True Then
    MATRIX_CROSS_CHECK_VECTOR_FUNC = 0
    For i = SROW To NROWS
        TEMP_SUM = 0
        For j = SCOLUMN To NCOLUMNS
            TEMP_SUM = TEMP_SUM + Abs(DATA_VECTOR(j, 1) - DATA_MATRIX(i, j))
        Next j
        If TEMP_SUM <= epsilon Then MATRIX_CROSS_CHECK_VECTOR_FUNC = i: Exit For
    Next i
Else
    GoTo ERROR_LABEL
End If

Exit Function
ERROR_LABEL:
MATRIX_CROSS_CHECK_VECTOR_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_CROSS_CHECK_VECTOR_FUNC
'DESCRIPTION   : Check for a reference value inside a vector

'LIBRARY       : MATRIX
'GROUP         : CHECK
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function VECTOR_CHECK_VALUE_FUNC(ByVal X_VAL As Variant, _
ByRef DATA_RNG As Variant)
  
Dim i As Long
Dim SROW As Long
Dim NROWS As Long
Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

VECTOR_CHECK_VALUE_FUNC = False

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then
  DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
End If
SROW = LBound(DATA_VECTOR, 1)
NROWS = UBound(DATA_VECTOR, 1)

If IS_1D_ARRAY_FUNC(DATA_VECTOR) = True Then
    For i = SROW To NROWS
        If X_VAL = DATA_VECTOR(i) Then
            VECTOR_CHECK_VALUE_FUNC = True
            Exit Function
        End If
    Next i
ElseIf IS_2D_ARRAY_FUNC(DATA_VECTOR) = True Then
    For i = SROW To NROWS
        If X_VAL = DATA_VECTOR(i, 1) Then
            VECTOR_CHECK_VALUE_FUNC = True
            Exit Function
        End If
    Next i
Else
    GoTo ERROR_LABEL
End If
  
Exit Function
ERROR_LABEL:
VECTOR_CHECK_VALUE_FUNC = False
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_FIND_ELEMENT_FUNC
'DESCRIPTION   : SEARCH AN ENTRY WITHIN ARRAY
'LIBRARY       : MATRIX
'GROUP         : CHECK
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_FIND_ELEMENT_FUNC(ByRef DATA_RNG As Variant, _
ByVal X_VAL As Variant, _
Optional ByVal SCOLUMN As Long = 1, _
Optional ByVal SROW As Long = 1, _
Optional ByVal VERSION As Integer = 0)

Dim MATCH_FLAG As Boolean
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
MATCH_FLAG = False

'---------------------------------------------------------------------------
Select Case VERSION
'---------------------------------------------------------------------------
Case 0 '--> FOR COLUMNS
'---------------------------------------------------------------------------
    Do Until SCOLUMN > UBound(DATA_MATRIX, 2)
        If DATA_MATRIX(SROW, SCOLUMN) Like X_VAL Then
            MATCH_FLAG = True
            Exit Do
        End If
        SCOLUMN = SCOLUMN + 1
    Loop
    If MATCH_FLAG = False Then
        MATRIX_FIND_ELEMENT_FUNC = 0
    Else
        MATRIX_FIND_ELEMENT_FUNC = SCOLUMN
    End If
'---------------------------------------------------------------------------
Case Else 'FOR ROWS
'---------------------------------------------------------------------------
    Do Until SROW > UBound(DATA_MATRIX, 1)
        If DATA_MATRIX(SROW, SCOLUMN) Like X_VAL Then
            MATCH_FLAG = True
            Exit Do
        End If
        SROW = SROW + 1
    Loop
    If MATCH_FLAG = False Then
        MATRIX_FIND_ELEMENT_FUNC = 0
    Else
        MATRIX_FIND_ELEMENT_FUNC = SROW
    End If
'---------------------------------------------------------------------------
End Select
'---------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
MATRIX_FIND_ELEMENT_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_CHECK_VALUE_FUNC
'DESCRIPTION   : Check for a reference value inside a matrix
'LIBRARY       : MATRIX
'GROUP         : CHECK
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_CHECK_VALUE_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal X_VAL As Variant = 0, _
Optional ByVal VERSION As Integer = 0)
  
Dim i As Long
Dim j As Long
  
Dim SROW As Long
Dim SCOLUMN As Long
  
Dim NROWS As Long
Dim NCOLUMNS As Long
  
Dim MATCH_FLAG As Variant
Dim DATA_MATRIX As Variant
    
On Error GoTo ERROR_LABEL
  
DATA_MATRIX = DATA_RNG
  
SROW = LBound(DATA_MATRIX, 1)
NROWS = UBound(DATA_MATRIX, 1)

SCOLUMN = LBound(DATA_MATRIX, 2)
NCOLUMNS = UBound(DATA_MATRIX, 2)

'---------------------------------------------------------------------------
Select Case VERSION
'---------------------------------------------------------------------------
Case 0
'---------------------------------------------------------------------------
    MATCH_FLAG = False
    For i = SROW To NROWS
        For j = SCOLUMN To NCOLUMNS
            If DATA_MATRIX(i, j) <> X_VAL Then
                MATCH_FLAG = True
                Exit For
            End If
        Next j
        If MATCH_FLAG = True Then: Exit For
    Next i
'---------------------------------------------------------------------------
Case Else
'---------------------------------------------------------------------------
    MATCH_FLAG = True
    For i = SROW To NROWS
        For j = SCOLUMN To NCOLUMNS
            If DATA_MATRIX(i, j) = X_VAL Then
                MATCH_FLAG = False
                Exit For
            End If
        Next j
        If MATCH_FLAG = False Then: Exit For
    Next i
'---------------------------------------------------------------------------
End Select
'---------------------------------------------------------------------------
  
MATRIX_CHECK_VALUE_FUNC = MATCH_FLAG
      
Exit Function
ERROR_LABEL:
MATRIX_CHECK_VALUE_FUNC = Err.number
End Function
