Attribute VB_Name = "MATRIX_LEONTIEF_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_LEONTIEF_FUNC
'DESCRIPTION   : Returns the inverse of Leontief matrix of the Input-Output
'Analysis Theory (Interdependence of industries). Parameter
'matrix is the interindustry exchange table (or IO-table). This
'table lists the value of the goods produced by each economic
'sector and how much of that output is used by each sector.
'Parameter vector is the total production vector.

'LIBRARY       : MATRIX
'GROUP         : LEONTIEF
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 14/02/2008
'************************************************************************************
'************************************************************************************

Function MATRIX_LEONTIEF_FUNC(ByRef MATRIX_RNG As Variant, _
ByRef VECTOR_RNG As Variant, _
Optional ByRef DEMAND_RNG As Variant, _
Optional ByVal OUTPUT As Integer = 2, _
Optional ByVal epsilon As Double = 2E-16)

Dim i As Long
Dim j As Long
Dim NSIZE As Long

Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant
Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = MATRIX_RNG 'Interindustry Exchange Matrix
DATA_VECTOR = VECTOR_RNG   'Total Production Vector
If UBound(DATA_VECTOR, 1) = 1 Then
    DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
End If
NSIZE = UBound(DATA_VECTOR, 1)
If UBound(DATA_MATRIX, 1) <> UBound(DATA_MATRIX, 2) Then GoTo ERROR_LABEL

If NSIZE <> UBound(DATA_MATRIX, 1) Then GoTo ERROR_LABEL

'build the consumption matrix
ReDim TEMP_MATRIX(1 To NSIZE, 1 To NSIZE)
For j = 1 To NSIZE
    For i = 1 To NSIZE
        TEMP_MATRIX(i, j) = DATA_MATRIX(i, j) / DATA_VECTOR(j, 1)
    Next i
Next j

For j = 1 To NSIZE
    For i = 1 To NSIZE
        TEMP_MATRIX(i, j) = -TEMP_MATRIX(i, j)
        If i = j Then TEMP_MATRIX(i, j) = 1 + TEMP_MATRIX(i, j)
    Next i
Next j
TEMP_MATRIX = MATRIX_LU_INVERSE_FUNC(TEMP_MATRIX)
'TEMP_MATRIX = MATRIX_GS_REDUCTION_PIVOT_FUNC(TEMP_MATRIX, , epsilon, 1)

Select Case OUTPUT
    Case 0
        MATRIX_LEONTIEF_FUNC = TEMP_MATRIX
    Case Else
        If IsArray(DEMAND_RNG) = True Then 'New Output
            DATA_VECTOR = DEMAND_RNG
            MATRIX_LEONTIEF_FUNC = MMULT_FUNC(TEMP_MATRIX, DATA_VECTOR, 70)
        Else: GoTo ERROR_LABEL
        End If
End Select

Exit Function
ERROR_LABEL:
MATRIX_LEONTIEF_FUNC = Err.number
End Function
