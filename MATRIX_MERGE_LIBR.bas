Attribute VB_Name = "MATRIX_MERGE_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : ARRAY_MERGE_FUNC
'DESCRIPTION   : Merge two arrays
'LIBRARY       : MATRIX
'GROUP         : MERGE
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function ARRAY_MERGE_FUNC(ByRef FIRST_RNG As Variant, _
ByRef SECOND_RNG As Variant, _
Optional ByVal epsilon As Double = 0.02)
  
Dim i As Long
Dim j As Long
Dim k As Long

Dim LONG_ARR As Variant
Dim SHORT_ARR As Variant

Dim CHECK_FLAG As Boolean
Dim MATCH_FLAG As Boolean
  
Dim FIRST_ARR As Variant
Dim SECOND_ARR As Variant

Dim TEMP_VALUE As Double

On Error GoTo ERROR_LABEL

FIRST_ARR = FIRST_RNG
SECOND_ARR = SECOND_RNG

If UBound(FIRST_ARR, 1) > UBound(SECOND_ARR, 1) Then
  LONG_ARR = FIRST_ARR
  SHORT_ARR = SECOND_ARR
Else
  LONG_ARR = SECOND_ARR
  SHORT_ARR = FIRST_ARR
End If

If IS_1D_ARRAY_FUNC(SHORT_ARR) = True Then
  For i = LBound(SHORT_ARR, 1) To UBound(SHORT_ARR, 1)
      TEMP_VALUE = SHORT_ARR(i)
      CHECK_FLAG = False
        For k = LBound(LONG_ARR) To UBound(LONG_ARR)
          If (Abs(TEMP_VALUE - LONG_ARR(k)) < epsilon) Then
            CHECK_FLAG = True
            GoTo 1983
          End If
        Next k
1983:
      If Not CHECK_FLAG Then
        ReDim TEMP_ARR(LBound(LONG_ARR) To UBound(LONG_ARR) + 1)
        MATCH_FLAG = False
        j = LBound(LONG_ARR)
        For k = LBound(LONG_ARR) To UBound(LONG_ARR, 1)
          If LONG_ARR(k) > TEMP_VALUE And MATCH_FLAG = False Then
            TEMP_ARR(j) = TEMP_VALUE
            j = j + 1
            MATCH_FLAG = True
          End If
          TEMP_ARR(j) = LONG_ARR(k)
          j = j + 1
        Next k
        If MATCH_FLAG = False Then: TEMP_ARR(j) = TEMP_VALUE
        LONG_ARR = TEMP_ARR
      End If
  Next i
ElseIf IS_2D_ARRAY_FUNC(SHORT_ARR) = True Then
    For i = LBound(SHORT_ARR, 1) To UBound(SHORT_ARR, 1)
      TEMP_VALUE = SHORT_ARR(i, 1)
      If Not VECTOR_CHECK_VALUE_FUNC(TEMP_VALUE, LONG_ARR) Then
        LONG_ARR = ARRAY_INSERT_VALUE_FUNC(LONG_ARR, TEMP_VALUE)
      End If
    Next i
End If

ARRAY_MERGE_FUNC = LONG_ARR

Exit Function
ERROR_LABEL:
ARRAY_MERGE_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : VECTOR_MERGE_FUNC
'DESCRIPTION   : Merge two vectors (with sorting option)
'LIBRARY       : MATRIX
'GROUP         : MERGE
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function VECTOR_MERGE_FUNC(ByRef ADATA_RNG As Variant, _
ByRef BDATA_RNG As Variant, _
Optional ByVal SORT_OPT As Integer = 1)

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long

Dim TEMP_VECTOR As Variant
Dim ADATA_VECTOR As Variant
Dim BDATA_VECTOR As Variant

Dim LONG_VECTOR As Variant
Dim SHORT_VECTOR As Variant

Dim TEMP_VALUE As Variant
Dim MATCH_FLAG As Boolean
Dim CHECK_FLAG As Boolean

On Error GoTo ERROR_LABEL

ADATA_VECTOR = ADATA_RNG
If UBound(ADATA_VECTOR, 1) = 1 Then
  ADATA_VECTOR = MATRIX_TRANSPOSE_FUNC(ADATA_VECTOR)
End If

BDATA_VECTOR = BDATA_RNG
If UBound(BDATA_VECTOR, 1) = 1 Then
  BDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(BDATA_VECTOR)
End If

If UBound(ADATA_VECTOR, 1) > UBound(BDATA_VECTOR, 1) Then
  LONG_VECTOR = ADATA_VECTOR
  SHORT_VECTOR = BDATA_VECTOR
Else
  LONG_VECTOR = BDATA_VECTOR
  SHORT_VECTOR = ADATA_VECTOR
End If

'-------------------------------------------------------------------------------------------------
For i = LBound(SHORT_VECTOR, 1) To UBound(SHORT_VECTOR, 1)
'-------------------------------------------------------------------------------------------------
    TEMP_VALUE = SHORT_VECTOR(i, 1)

'---------------------------Check if an element is in the vector---------------------------
    MATCH_FLAG = False
    For j = LBound(LONG_VECTOR, 1) To UBound(LONG_VECTOR, 1)
        If TEMP_VALUE = LONG_VECTOR(j, 1) Then
            MATCH_FLAG = True
                Exit For
        End If
    Next j

'----------------Inserts a value in a sorted array in the sorted position-----------------

    If Not MATCH_FLAG Then
      ReDim TEMP_VECTOR(LBound(LONG_VECTOR, 1) To _
                        UBound(LONG_VECTOR, 1) + 1, 1 To 1)
      CHECK_FLAG = False
      h = LBound(LONG_VECTOR, 1)
      For k = LBound(LONG_VECTOR, 1) To UBound(LONG_VECTOR, 1)
            If LONG_VECTOR(k, 1) > TEMP_VALUE And CHECK_FLAG = False Then
                  TEMP_VECTOR(h, 1) = TEMP_VALUE
                  h = h + 1
                  CHECK_FLAG = True
            End If
            TEMP_VECTOR(h, 1) = LONG_VECTOR(k, 1)
            h = h + 1
      Next k
        
      If CHECK_FLAG = False Then: TEMP_VECTOR(h, 1) = TEMP_VALUE
      LONG_VECTOR = TEMP_VECTOR
    End If
'-------------------------------------------------------------------------------------------------
Next i
'-------------------------------------------------------------------------------------------------

If SORT_OPT <> 0 Then: LONG_VECTOR = MATRIX_QUICK_SORT_FUNC(LONG_VECTOR, 1, 1)
VECTOR_MERGE_FUNC = LONG_VECTOR
  
Exit Function
ERROR_LABEL:
VECTOR_MERGE_FUNC = Err.number
End Function
