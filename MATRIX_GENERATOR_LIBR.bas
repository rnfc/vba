Attribute VB_Name = "MATRIX_GENERATOR_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_GENERATOR_FUNC
'DESCRIPTION   : Returns a (NROWS x NCOLUMNS) Matrix
'LIBRARY       : MATRIX
'GROUP         : GENERATOR
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_GENERATOR_FUNC(ByVal NROWS As Long, _
ByVal NCOLUMNS As Long, _
Optional ByVal REF_VALUE As Variant = "")

Dim i As Long
Dim j As Long
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL
    
ReDim TEMP_MATRIX(1 To NROWS, 1 To NCOLUMNS)
For j = 1 To NCOLUMNS
    For i = 1 To NROWS
        TEMP_MATRIX(i, j) = REF_VALUE
    Next i
Next j

MATRIX_GENERATOR_FUNC = TEMP_MATRIX
    
Exit Function
ERROR_LABEL:
MATRIX_GENERATOR_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_IDENTITY_FUNC
'DESCRIPTION   : Returns the (n x n) Identity Matrix
'LIBRARY       : MATRIX
'GROUP         : GENERATOR
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_IDENTITY_FUNC(ByVal NSIZE As Long)

Dim i As Long
Dim TEMP_MATRIX As Variant
    
On Error GoTo ERROR_LABEL
    
ReDim TEMP_MATRIX(1 To NSIZE, 1 To NSIZE)
For i = 1 To NSIZE
    TEMP_MATRIX(i, i) = 1
Next i
    
MATRIX_IDENTITY_FUNC = TEMP_MATRIX
    
Exit Function
ERROR_LABEL:
MATRIX_IDENTITY_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : VECTOR_IDENTITY_FUNC
'DESCRIPTION   : Returns the (1 x 1) Identity Matrix
'LIBRARY       : MATRIX
'GROUP         : GENERATOR
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function VECTOR_IDENTITY_FUNC(ByVal NROWS As Long)
    
Dim i As Long
Dim TEMP_VECTOR As Variant
  
ReDim TEMP_VECTOR(1 To NROWS, 1 To 1)
For i = 1 To NROWS
    TEMP_VECTOR(i, 1) = 1
Next i
    
VECTOR_IDENTITY_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
VECTOR_IDENTITY_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : RNG_MATRIX_GENERATOR_FUNC
'DESCRIPTION   : Generate a matrix (n x n) in Excel
'LIBRARY       : MATRIX
'GROUP         : GENERATOR
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function RNG_MATRIX_GENERATOR_FUNC(ByRef DST_RNG As Excel.Range, _
ByVal NSIZE As Long, _
Optional ByVal FORMULA_STR As String = "=Rand()*100+1", _
Optional ByVal HEADER_STR As String = "XXXX - ")
    
Dim i As Long
Dim j As Long

Dim TEMP_RNG As Excel.Range
Dim DATA_RNG As Excel.Range

On Error GoTo ERROR_LABEL

RNG_MATRIX_GENERATOR_FUNC = False
Set TEMP_RNG = DST_RNG.Cells(1, 1)

With TEMP_RNG
    Set DATA_RNG = Range(.Offset(NSIZE, 1), .Offset(1, NSIZE))
    DATA_RNG.ClearContents
    For i = 1 To NSIZE
        .Offset(0, i).value = HEADER_STR & CStr(i)
    Next i
    For j = 1 To NSIZE
        .Offset(j, 0).formula = "=offset(" & TEMP_RNG.Address _
            & ",0," & CStr(j) & ")"
    Next j
End With

With DATA_RNG
     For i = 1 To NSIZE - 1
        For j = i + 1 To NSIZE
            .Cells(i, j) = FORMULA_STR
        Next
    Next
    For i = 2 To NSIZE
        For j = 1 To i - 1
            .Cells(i, j).formula = "=offset(" & TEMP_RNG.Address _
                & "," & CStr(j) & "," & CStr(i) & ")"
        Next j
    Next i
End With

TEMP_RNG.Select
RNG_MATRIX_GENERATOR_FUNC = True
    
Exit Function
ERROR_LABEL:
RNG_MATRIX_GENERATOR_FUNC = False
End Function
