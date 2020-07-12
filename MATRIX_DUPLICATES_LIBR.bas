Attribute VB_Name = "MATRIX_DUPLICATES_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : ARRAY_REMOVE_DUPLICATES_FUNC
'DESCRIPTION   : Remove duplicates rows from a vector
'LIBRARY       : MATRIX
'GROUP         : DUPLICATES
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function ARRAY_REMOVE_DUPLICATES_FUNC(ByRef DATA_VECTOR As Variant, _
Optional ByVal VERSION As Integer = 0)
'ByRef DATA_RNG As Variant

'VERSION 0 --> Recognize single space
'VERSION 1 --> Do not recognize single space

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim TEMP_STR As String
'Dim DATA_VECTOR As Variant

'If you want to use this function inside excel instead of calling the main routine
'you must change the input DATA_VECTOR for DATA_RNG and declare DATA_VECTOR again
'inside the function. This is a trick to save memoory, so we don't have two
'arrays declared.

Dim TEMP_VECTOR As Variant
Dim COLLECTION_OBJ As New Collection

If IsArray(DATA_VECTOR) = False Then: GoTo ERROR_LABEL

On Error Resume Next

NROWS = UBound(DATA_VECTOR)
If IS_2D_ARRAY_FUNC(DATA_VECTOR) Then
    For i = 1 To NROWS
        TEMP_STR = DATA_VECTOR(i, 1)
        If VERSION = 0 Then
            TEMP_STR = Trim(REMOVE_EXTRA_SPACES_FUNC(TEMP_STR))
        Else
            TEMP_STR = Trim(Replace(REMOVE_EXTRA_SPACES_FUNC(TEMP_STR), " ", ""))
        End If
        Call COLLECTION_OBJ.Add(CStr(i), TEMP_STR)
        DATA_VECTOR(i, 1) = TEMP_STR
        If Err.number <> 0 Then: Err.number = 0 'Found Duplicate
    Next i
    NROWS = COLLECTION_OBJ.COUNT
    ReDim TEMP_VECTOR(1 To NROWS, 1 To 1)
    For i = 1 To NROWS
        j = CLng(COLLECTION_OBJ.Item(i))
        TEMP_VECTOR(i, 1) = DATA_VECTOR(j, 1)
    Next i
Else
    For i = 1 To NROWS
        TEMP_STR = DATA_VECTOR(i)
        If VERSION = 0 Then
            TEMP_STR = Trim(REMOVE_EXTRA_SPACES_FUNC(TEMP_STR))
        Else
            TEMP_STR = Trim(Replace(REMOVE_EXTRA_SPACES_FUNC(TEMP_STR), " ", ""))
        End If
        Call COLLECTION_OBJ.Add(CStr(i), TEMP_STR)
        DATA_VECTOR(i) = TEMP_STR
        If Err.number <> 0 Then: Err.number = 0 'Found Duplicate
    Next i
    NROWS = COLLECTION_OBJ.COUNT
    ReDim TEMP_VECTOR(1 To NROWS)
    For i = 1 To NROWS
        j = CLng(COLLECTION_OBJ.Item(i))
        TEMP_VECTOR(i) = DATA_VECTOR(j)
    Next i
End If

ARRAY_REMOVE_DUPLICATES_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
ARRAY_REMOVE_DUPLICATES_FUNC = Err.number
End Function
