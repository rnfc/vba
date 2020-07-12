Attribute VB_Name = "OPTIM_MULTVAR_OBJ_LIBR"

'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
Option Explicit     'Requires that all variables to be declared explicitly.
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : MULTVAR_CALL_OBJ_FUNC
'DESCRIPTION   : Load Objective Function
'LIBRARY       : OPTIMIZATION
'GROUP         : MULTVAR_OBJ
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function MULTVAR_CALL_OBJ_FUNC(ByVal FUNC_NAME_STR As Variant, _
ByRef PARAM_RNG As Variant, _
Optional ByRef SCALE_RNG As Variant, _
Optional ByVal MIN_FLAG As Boolean = True)

Dim i As Long
Dim j As Long
Dim NROWS As Long
Dim NCOLUMNS As Long

Dim YTEMP_VAL As Double
Dim TEMP_FACTOR As Double

Dim YTEMP_VECTOR As Variant
Dim XTEMP_VECTOR As Variant

Dim PARAM_VECTOR As Variant
Dim SCALE_VECTOR As Variant

On Error GoTo ERROR_LABEL

If MIN_FLAG Then
    TEMP_FACTOR = 1
Else
    TEMP_FACTOR = -1
End If

PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)
NROWS = UBound(PARAM_VECTOR, 1)  'number of points
NCOLUMNS = UBound(PARAM_VECTOR, 2)  'number of  variables

If IsArray(SCALE_RNG) = True Then
    SCALE_VECTOR = SCALE_RNG
    If UBound(SCALE_VECTOR, 1) = 1 Then: SCALE_VECTOR = MATRIX_TRANSPOSE_FUNC(SCALE_VECTOR)
Else
    ReDim SCALE_VECTOR(1 To NROWS, 1 To 1)
    For i = 1 To NROWS
        SCALE_VECTOR(i, 1) = 1
    Next i
End If

'---------------------------------------------------------------------------
If NCOLUMNS > 1 Then
'---------------------------------------------------------------------------
    ReDim YTEMP_VECTOR(1 To NROWS, 1 To 1)
    ReDim XTEMP_VECTOR(1 To NCOLUMNS, 1 To 1)

    For i = 1 To NROWS
        For j = 1 To NCOLUMNS
            XTEMP_VECTOR(j, 1) = SCALE_VECTOR(j, 1) * PARAM_VECTOR(i, j)
        Next j
        YTEMP_VAL = Excel.Application.Run(FUNC_NAME_STR, XTEMP_VECTOR)
        YTEMP_VAL = YTEMP_VAL * TEMP_FACTOR
    
        YTEMP_VECTOR(i, 1) = YTEMP_VAL
    Next i

    MULTVAR_CALL_OBJ_FUNC = YTEMP_VECTOR

'---------------------------------------------------------------------------
Else
'---------------------------------------------------------------------------
    
    For i = 1 To NROWS
        PARAM_VECTOR(i, 1) = SCALE_VECTOR(i, 1) * PARAM_VECTOR(i, 1)
    Next i
        
    YTEMP_VAL = Excel.Application.Run(FUNC_NAME_STR, PARAM_VECTOR)
    YTEMP_VAL = YTEMP_VAL * TEMP_FACTOR
    
    MULTVAR_CALL_OBJ_FUNC = YTEMP_VAL

'---------------------------------------------------------------------------
End If
'---------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
MULTVAR_CALL_OBJ_FUNC = Err.number
End Function
