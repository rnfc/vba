Attribute VB_Name = "OPTIM_MULTVAR_GRAD_LIBR"

'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
Option Explicit     'Requires that all variables to be declared explicitly.
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------


'************************************************************************************
'************************************************************************************
'FUNCTION      : MULTVAR_CALL_GRAD_FUNC
'DESCRIPTION   : Load Gradients
'LIBRARY       : OPTIMIZATION
'GROUP         : MULTVAR_GRAD
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function MULTVAR_CALL_GRAD_FUNC(ByVal GRAD_STR_NAME As String, _
ByRef PARAM_RNG As Variant, _
Optional ByRef SCALE_RNG As Variant, _
Optional ByVal MIN_FLAG As Boolean = True)

Dim ii As Long
Dim NROWS As Long

Dim TEMP_FACTOR As Double

Dim GRAD_VECTOR As Variant
Dim SCALE_VECTOR As Variant
Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL

If MIN_FLAG Then
    TEMP_FACTOR = 1
Else
    TEMP_FACTOR = -1
End If

PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)
NROWS = UBound(PARAM_VECTOR, 1)

If IsArray(SCALE_RNG) = True Then
    SCALE_VECTOR = SCALE_RNG
    If UBound(SCALE_VECTOR, 1) = 1 Then: _
        SCALE_VECTOR = MATRIX_TRANSPOSE_FUNC(SCALE_VECTOR)
Else
    ReDim SCALE_VECTOR(1 To NROWS, 1 To 1)
    For ii = 1 To NROWS
        SCALE_VECTOR(ii, 1) = 1
    Next ii
End If

For ii = 1 To NROWS
    PARAM_VECTOR(ii, 1) = PARAM_VECTOR(ii, 1) * SCALE_VECTOR(ii, 1)
Next ii
    
GRAD_VECTOR = Excel.Application.Run(GRAD_STR_NAME, PARAM_VECTOR)
NROWS = UBound(GRAD_VECTOR, 1)
    
For ii = 1 To NROWS
    GRAD_VECTOR(ii, 1) = TEMP_FACTOR * GRAD_VECTOR(ii, 1)
Next ii

MULTVAR_CALL_GRAD_FUNC = GRAD_VECTOR

Exit Function
ERROR_LABEL:
    MULTVAR_CALL_GRAD_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MULTVAR_CALL_POINT_GRAD_FUNC
'DESCRIPTION   : Evaluate gradient function
'LIBRARY       : OPTIMIZATION
'GROUP         : MULTVAR_GRAD
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function MULTVAR_CALL_POINT_GRAD_FUNC(ByVal FUNC_NAME_STR As String, _
ByVal GRAD_STR_NAME As String, _
ByRef PARAM_RNG As Variant, _
ByVal kk As Long, _
Optional ByRef SCALE_RNG As Variant, _
Optional ByVal MIN_FLAG As Boolean = True)

Dim ii As Long
Dim jj As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_FACTOR As Double

Dim XTEMP_VECTOR As Variant
Dim YTEMP_VECTOR As Variant

Dim GRAD_VECTOR As Variant
Dim PARAM_MATRIX As Variant
Dim SCALE_VECTOR As Variant

On Error GoTo ERROR_LABEL

If MIN_FLAG Then
    TEMP_FACTOR = 1
Else
    TEMP_FACTOR = -1
End If

PARAM_MATRIX = PARAM_RNG
If UBound(PARAM_MATRIX, 1) = 1 Then: PARAM_MATRIX = MATRIX_TRANSPOSE_FUNC(PARAM_MATRIX)
NROWS = UBound(PARAM_MATRIX, 1)  'number of points
NCOLUMNS = UBound(PARAM_MATRIX, 2)  'number of  variables

If IsArray(SCALE_RNG) = True Then
    SCALE_VECTOR = SCALE_RNG
    If UBound(SCALE_VECTOR, 1) = 1 Then: SCALE_VECTOR = MATRIX_TRANSPOSE_FUNC(SCALE_VECTOR)
Else
    ReDim SCALE_VECTOR(1 To NCOLUMNS, 1 To 1)
    For ii = 1 To NCOLUMNS
        SCALE_VECTOR(ii, 1) = 1
    Next ii
End If

ReDim GRAD_VECTOR(1 To NROWS, 1 To 1)
ReDim XTEMP_VECTOR(1 To NCOLUMNS, 1 To 1)

For ii = 1 To NROWS
    For jj = 1 To NCOLUMNS
        XTEMP_VECTOR(jj, 1) = SCALE_VECTOR(jj, 1) * PARAM_MATRIX(ii, jj)
    Next jj
    YTEMP_VECTOR = Excel.Application.Run(GRAD_STR_NAME, XTEMP_VECTOR)
    
    GRAD_VECTOR(ii, 1) = TEMP_FACTOR * YTEMP_VECTOR(kk, 1)
Next ii

MULTVAR_CALL_POINT_GRAD_FUNC = GRAD_VECTOR

Exit Function
ERROR_LABEL:
    MULTVAR_CALL_POINT_GRAD_FUNC = Err.number
End Function
