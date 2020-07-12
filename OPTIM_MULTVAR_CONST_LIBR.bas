Attribute VB_Name = "OPTIM_MULTVAR_CONST_LIBR"
'// PERFECT

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
'FUNCTION      : MULTVAR_LOAD_CONST_FUNC
'DESCRIPTION   : Load gradient function
'LIBRARY       : OPTIMIZATION
'GROUP         : MULTVAR_CONST
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function MULTVAR_LOAD_CONST_FUNC(ByRef CONST_RNG As Variant, _
Optional ByVal VERSION As Integer = 0)

Dim i As Long
Dim NROWS As Long
Dim NCOLUMNS As Long

Dim CONST_DATA As Variant
Dim CONST_BOX As Variant

On Error GoTo ERROR_LABEL

CONST_DATA = CONST_RNG
NROWS = UBound(CONST_DATA, 1)
NCOLUMNS = UBound(CONST_DATA, 2)

'------------------------------------------------------------------------------
If VERSION = 0 Then 'vertical box
'------------------------------------------------------------------------------
    ReDim CONST_BOX(1 To NROWS, 1 To 2)
    For i = 1 To NROWS
        CONST_BOX(i, 1) = CONST_DATA(i, 1)       'Xmin
        CONST_BOX(i, 2) = CONST_DATA(i, 2)       'Xmax
    Next i
'------------------------------------------------------------------------------
Else
'------------------------------------------------------------------------------
    ReDim CONST_BOX(1 To NCOLUMNS, 1 To 2)
    For i = 1 To NCOLUMNS
        CONST_BOX(i, 1) = CONST_DATA(1, i)       'Xmin
        CONST_BOX(i, 2) = CONST_DATA(2, i)       'Xmax
    Next i
'------------------------------------------------------------------------------
End If
'------------------------------------------------------------------------------

MULTVAR_LOAD_CONST_FUNC = CONST_BOX

Exit Function
ERROR_LABEL:
    MULTVAR_LOAD_CONST_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MULTVAR_SCALE_CONST_FUNC
'DESCRIPTION   : Rescale the variables with an adapt factor
'LIBRARY       : OPTIMIZATION
'GROUP         : MULTVAR_CONST
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function MULTVAR_SCALE_CONST_FUNC(ByRef CONST_RNG As Variant)

Dim i As Long
Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_VALUE As Double

Dim CONST_BOX As Variant
Dim SCALE_VECTOR As Variant

On Error GoTo ERROR_LABEL

CONST_BOX = CONST_RNG
NROWS = UBound(CONST_BOX, 1)
NCOLUMNS = UBound(CONST_BOX, 2)

ReDim SCALE_VECTOR(1 To NROWS, 1 To 1)
'--------------------------------------------------------------------------
If NCOLUMNS = 2 Then
'--------------------------------------------------------------------------
    For i = 1 To NROWS
        If (CONST_BOX(i, 2) - CONST_BOX(i, 1)) <> 0 Then
              TEMP_VALUE = _
                Int(Log(Abs((CONST_BOX(i, 2) - CONST_BOX(i, 1)))) / Log(10#))
        Else
              TEMP_VALUE = 0
        End If
        SCALE_VECTOR(i, 1) = 10 ^ TEMP_VALUE
        CONST_BOX(i, 1) = CONST_BOX(i, 1) / SCALE_VECTOR(i, 1)
        CONST_BOX(i, 2) = CONST_BOX(i, 2) / SCALE_VECTOR(i, 1)
    Next i
'--------------------------------------------------------------------------
Else
'--------------------------------------------------------------------------
    For i = 1 To NROWS
        If CONST_BOX(i, 1) <> 0 Then
              TEMP_VALUE = Int(Log(Abs(CONST_BOX(i, 1))) / Log(10#))
        Else
              TEMP_VALUE = 0
        End If
        SCALE_VECTOR(i, 1) = 10 ^ TEMP_VALUE
        CONST_BOX(i, 1) = CONST_BOX(i, 1) / SCALE_VECTOR(i, 1)
    Next i
'--------------------------------------------------------------------------
End If
'--------------------------------------------------------------------------

MULTVAR_SCALE_CONST_FUNC = Array(CONST_BOX, SCALE_VECTOR)

Exit Function
ERROR_LABEL:
    MULTVAR_SCALE_CONST_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MULTVAR_CONST_BOUND_FUNC
'DESCRIPTION   : Find the protection point on the Constraint Box
'LIBRARY       : OPTIMIZATION
'GROUP         : MULTVAR_CONST
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function MULTVAR_CONST_BOUND_FUNC(ByRef CONST_RNG As Variant, _
ByRef DATA_RNG As Variant, _
ByRef PARAM_RNG As Variant)

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long

Dim NROWS As Long
Dim COUNTER As Long

Dim TEMP_NORM As Double
Dim TEMP_VALUE As Double
Dim TEMP_DELTA As Double

Dim CONST_BOX As Variant
Dim DATA_VECTOR As Variant
Dim PARAM_VECTOR As Variant

Dim XTEMP_VECTOR As Variant
Dim YTEMP_VECTOR As Variant

On Error GoTo ERROR_LABEL

CONST_BOX = CONST_RNG

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then: DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)

PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)

NROWS = UBound(PARAM_VECTOR)
YTEMP_VECTOR = DATA_VECTOR

TEMP_NORM = 0 'return the Euclidean norm of a vector
For i = 1 To UBound(YTEMP_VECTOR)
    TEMP_NORM = TEMP_NORM + YTEMP_VECTOR(i, 1) ^ 2
Next i
TEMP_VALUE = Sqr(TEMP_NORM)

If TEMP_VALUE = 0 Then
    ReDim XTEMP_VECTOR(1 To NROWS, 1 To 1)
    For i = 1 To NROWS
        XTEMP_VECTOR(i, 1) = PARAM_VECTOR(i, 1)
    Next i
Else
    For i = 1 To NROWS
        YTEMP_VECTOR(i, 1) = YTEMP_VECTOR(i, 1) / TEMP_VALUE
    Next i
    'finder algorithm begin
    ReDim XTEMP_VECTOR(1 To NROWS, 1 To 1)
    k = 1
    COUNTER = 0
    Do
        If YTEMP_VECTOR(k, 1) = 0 Then
            k = k + 1
        Else
            XTEMP_VECTOR(k, 1) = CONST_BOX(k, 1)
            TEMP_DELTA = (XTEMP_VECTOR(k, 1) - PARAM_VECTOR(k, 1)) / YTEMP_VECTOR(k, 1)
            If TEMP_DELTA <= 0 Then
                'change direction
                XTEMP_VECTOR(k, 1) = CONST_BOX(k, 2)
                TEMP_DELTA = (XTEMP_VECTOR(k, 1) - PARAM_VECTOR(k, 1)) / _
                            YTEMP_VECTOR(k, 1)
            End If
            'compute the new point
            For j = 1 To NROWS
                If j <> k Then _
                    XTEMP_VECTOR(j, 1) = PARAM_VECTOR(j, 1) + _
                                    TEMP_DELTA * YTEMP_VECTOR(j, 1)
            Next j
            h = 0
            For j = 1 To UBound(XTEMP_VECTOR) 'return the variable entry
            'having the constrain violated otherwise return 0
                If XTEMP_VECTOR(j, 1) < CONST_BOX(j, 1) Or _
                   XTEMP_VECTOR(j, 1) > CONST_BOX(j, 2) Then
                    h = j
                    Exit For
                End If
            Next j
            k = h 'check constrain
        End If
        COUNTER = COUNTER + 1
    Loop Until k = 0 Or COUNTER > NROWS
End If

MULTVAR_CONST_BOUND_FUNC = XTEMP_VECTOR

Exit Function
ERROR_LABEL:
MULTVAR_CONST_BOUND_FUNC = Err.number
End Function
