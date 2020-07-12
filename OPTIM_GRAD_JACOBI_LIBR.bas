Attribute VB_Name = "OPTIM_GRAD_JACOBI_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
                    
'// PERFECT


'************************************************************************************
'************************************************************************************
'FUNCTION      : JACOBI_MATRIX_VALID_FUNC
'DESCRIPTION   : Validate Jacobian Matrix
'LIBRARY       : OPTIMIZATION
'GROUP         : GRAD_JACOBI
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function JACOBI_MATRIX_VALID_FUNC(ByVal FUNC_NAME_STR As String, _
ByVal GRAD_STR_NAME As String, _
ByRef XDATA_RNG As Variant, _
ByRef PARAM_RNG As Variant, _
Optional ByVal MIN_FLAG As Boolean = True, _
Optional ByVal epsilon As Double = 10 ^ -5)

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_VAL As Double
Dim FACTOR_VAL As Double

Dim XDATA_MATRIX As Variant
Dim JACOBI_MATRIX As Variant
Dim GRAD_MATRIX As Variant
Dim PARAM_VECTOR As Variant

Dim LAMBDA As Double

On Error GoTo ERROR_LABEL

JACOBI_MATRIX_VALID_FUNC = True

LAMBDA = 1000
If MIN_FLAG = True Then
    FACTOR_VAL = 1
Else
    FACTOR_VAL = -1
End If

XDATA_MATRIX = XDATA_RNG
PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)

'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
GRAD_MATRIX = Excel.Application.Run(GRAD_STR_NAME, XDATA_MATRIX, PARAM_VECTOR)
NROWS = UBound(GRAD_MATRIX, 1)
NCOLUMNS = UBound(GRAD_MATRIX, 2)

For j = 1 To NCOLUMNS
    For i = 1 To NROWS
        GRAD_MATRIX(i, j) = GRAD_MATRIX(i, j) * FACTOR_VAL
    Next i
Next j
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------

JACOBI_MATRIX = JACOBI_MATRIX_FUNC(FUNC_NAME_STR, XDATA_MATRIX, PARAM_VECTOR, epsilon)
NROWS = UBound(JACOBI_MATRIX, 1)
NCOLUMNS = UBound(JACOBI_MATRIX, 2)

For j = 1 To NCOLUMNS
    For i = 1 To NROWS
        JACOBI_MATRIX(i, j) = JACOBI_MATRIX(i, j) * FACTOR_VAL
    Next i
Next j

'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
    
TEMP_VAL = 0
For j = 1 To NCOLUMNS
    For i = 1 To NROWS
        TEMP_VAL = Abs(GRAD_MATRIX(i, j) - JACOBI_MATRIX(i, j))
        If Abs(GRAD_MATRIX(i, j)) > epsilon Then TEMP_VAL = TEMP_VAL / Abs(GRAD_MATRIX(i, j))
        If TEMP_VAL > LAMBDA * epsilon Then
            JACOBI_MATRIX_VALID_FUNC = False 'Derivatives: dubious accuracy. Check the formula
            Exit Function
        End If
    Next i
Next j

Exit Function
ERROR_LABEL:
JACOBI_MATRIX_VALID_FUNC = False
End Function

'// PERFECT


'************************************************************************************
'************************************************************************************
'FUNCTION      : JACOBI_MATRIX_FUNC
'DESCRIPTION   : Returns the Jacobian Matrix
'LIBRARY       : OPTIMIZATION
'GROUP         : GRAD_JACOBI
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function JACOBI_MATRIX_FUNC(ByVal FUNC_NAME_STR As String, _
ByRef DATA_RNG As Variant, _
ByRef PARAM_RNG As Variant, _
Optional ByRef EPS_RNG As Variant = 0.001)
  
'FUNC_NAME_STR is the name of the function
'DATA_MATRIX is vector of independent variables
'PARAM_VECTOR is vector of parameter values
'DELTA_VECTOR is vector of fractional increments for each parameter
  
Dim i As Long
Dim j As Long

Dim NSIZE As Long
Dim NROWS As Long

Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant
Dim PARAM_VECTOR As Variant
Dim EPSILON_VECTOR As Variant

Dim XTEMP_VECTOR As Variant
Dim XDATA_VECTOR As Variant
Dim YDATA_VECTOR As Variant

Dim TEMP1_VECTOR As Variant
Dim TEMP2_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
If UBound(DATA_MATRIX, 1) = 1 Then: DATA_MATRIX = MATRIX_TRANSPOSE_FUNC(DATA_MATRIX)
NROWS = UBound(DATA_MATRIX, 1)

PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)
NSIZE = UBound(PARAM_VECTOR, 1)

If IsArray(EPS_RNG) = False Then
    ReDim EPSILON_VECTOR(1 To NSIZE, 1 To 1)
    For i = 1 To NSIZE
          EPSILON_VECTOR(i, 1) = EPS_RNG
    Next i
Else
    EPSILON_VECTOR = EPS_RNG
    If UBound(EPSILON_VECTOR, 1) = 1 Then: _
      EPSILON_VECTOR = MATRIX_TRANSPOSE_FUNC(EPSILON_VECTOR)
End If

XDATA_VECTOR = PARAM_VECTOR
YDATA_VECTOR = Excel.Application.Run(FUNC_NAME_STR, DATA_MATRIX, PARAM_VECTOR)

ReDim XTEMP_VECTOR(1 To NSIZE, 1 To 1)
ReDim TEMP_MATRIX(1 To NROWS, 1 To NSIZE)

For i = 1 To NSIZE 'loop over each parameter
  'calculate increment for parameter
    If PARAM_VECTOR(i, 1) = 0 Then
        XTEMP_VECTOR(i, 1) = EPSILON_VECTOR(i, 1)
    Else
        XTEMP_VECTOR(i, 1) = EPSILON_VECTOR(i, 1) * PARAM_VECTOR(i, 1)
    End If
  
    PARAM_VECTOR(i, 1) = XDATA_VECTOR(i, 1) + XTEMP_VECTOR(i, 1)
    
    If XTEMP_VECTOR(i, 1) <> 0 Then
        TEMP1_VECTOR = Excel.Application.Run(FUNC_NAME_STR, DATA_MATRIX, PARAM_VECTOR)
        '------------------------------------------------------------------------------------
        If EPSILON_VECTOR(i, 1) < 0 Then 'use forward difference
        '------------------------------------------------------------------------------------
            For j = 1 To NROWS
                TEMP_MATRIX(j, i) = (TEMP1_VECTOR(j, 1) - YDATA_VECTOR(j, 1)) / XTEMP_VECTOR(i, 1)
            Next j
        '------------------------------------------------------------------------------------
        Else 'use central difference
        '------------------------------------------------------------------------------------
            PARAM_VECTOR(i, 1) = XDATA_VECTOR(i, 1) - XTEMP_VECTOR(i, 1)
            TEMP2_VECTOR = Excel.Application.Run(FUNC_NAME_STR, DATA_MATRIX, PARAM_VECTOR)
            For j = 1 To NROWS
                TEMP_MATRIX(j, i) = (TEMP1_VECTOR(j, 1) - TEMP2_VECTOR(j, 1)) / (2 * XTEMP_VECTOR(i, 1))
            Next j
        '------------------------------------------------------------------------------------
        End If
        '------------------------------------------------------------------------------------
    End If
    '------------------------------------------------------------------------------------
    'this is a partial derivative, reset the previous value
    PARAM_VECTOR(i, 1) = XDATA_VECTOR(i, 1)
Next i

JACOBI_MATRIX_FUNC = TEMP_MATRIX
  
Exit Function
ERROR_LABEL:
JACOBI_MATRIX_FUNC = Err.number
End Function

'// PERFECT

'************************************************************************************
'************************************************************************************
'FUNCTION      : JACOBI_CENTRAL_FUNC
'DESCRIPTION   : Compute the Jacobian with the central step FD formula
'LIBRARY       : OPTIMIZATION
'GROUP         : GRAD_JACOBI
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function JACOBI_CENTRAL_FUNC(ByVal FUNC_NAME_STR As String, _
ByRef PARAM_RNG As Variant, _
Optional ByVal epsilon As Double = 10 ^ -5)

Dim i As Long
Dim j As Long

Dim NSIZE As Long
Dim NROWS As Long

Dim GRAD_VAL As Double
Dim TEMP_ARR As Variant
Dim TEMP_MATRIX As Variant

Dim TEMP1_VECTOR As Variant
Dim TEMP2_VECTOR As Variant

Dim XTEMP_VECTOR As Variant

Dim XDATA_VECTOR As Variant
Dim YDATA_VECTOR As Variant
Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL

PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)
NSIZE = UBound(PARAM_VECTOR, 1)

XDATA_VECTOR = PARAM_VECTOR
ReDim XTEMP_VECTOR(1 To NSIZE, 1 To 1)

TEMP_ARR = Excel.Application.Run(FUNC_NAME_STR, PARAM_VECTOR)
If IsArray(TEMP_ARR) = False Then
    ReDim YDATA_VECTOR(1 To 1, 1 To 1)
    YDATA_VECTOR(1, 1) = TEMP_ARR
Else
    YDATA_VECTOR = TEMP_ARR
    If UBound(YDATA_VECTOR, 1) = 1 Then: YDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(YDATA_VECTOR)
End If

NROWS = UBound(YDATA_VECTOR, 1)

'-------------------------------------------------------------------------------
ReDim TEMP_MATRIX(1 To NROWS, 1 To NSIZE)
'-------------------------------------------------------------------------------

For j = 1 To NSIZE 'forward step
    GRAD_VAL = epsilon * XDATA_VECTOR(j, 1)
    If GRAD_VAL < epsilon Then GRAD_VAL = epsilon
    For i = 1 To NSIZE
        If i = j Then
            XTEMP_VECTOR(i, 1) = XDATA_VECTOR(i, 1) + GRAD_VAL
        Else
            XTEMP_VECTOR(i, 1) = XDATA_VECTOR(i, 1)
        End If
    Next i
    
    TEMP_ARR = Excel.Application.Run(FUNC_NAME_STR, XTEMP_VECTOR)
    If IsArray(TEMP_ARR) = False Then
        ReDim TEMP1_VECTOR(1 To 1, 1 To 1)
        TEMP1_VECTOR(1, 1) = TEMP_ARR
    Else
        TEMP1_VECTOR = TEMP_ARR
        If UBound(TEMP1_VECTOR, 1) = 1 Then: TEMP1_VECTOR = MATRIX_TRANSPOSE_FUNC(TEMP1_VECTOR)
    End If
    
    GRAD_VAL = epsilon * XDATA_VECTOR(j, 1)
    If GRAD_VAL < epsilon Then GRAD_VAL = epsilon
    For i = 1 To NSIZE
        If i = j Then
            XTEMP_VECTOR(i, 1) = XDATA_VECTOR(i, 1) - GRAD_VAL
        Else
            XTEMP_VECTOR(i, 1) = XDATA_VECTOR(i, 1)
        End If
    Next i
    
    TEMP_ARR = Excel.Application.Run(FUNC_NAME_STR, XTEMP_VECTOR)
    If IsArray(TEMP_ARR) = False Then
        ReDim TEMP2_VECTOR(1 To 1, 1 To 1)
        TEMP2_VECTOR(1, 1) = TEMP_ARR
    Else
        TEMP2_VECTOR = TEMP_ARR
        If UBound(TEMP2_VECTOR, 1) = 1 Then: TEMP2_VECTOR = MATRIX_TRANSPOSE_FUNC(TEMP2_VECTOR)
    End If
    'FD central formula
    For i = 1 To NROWS
        TEMP_MATRIX(i, j) = (TEMP1_VECTOR(i, 1) - TEMP2_VECTOR(i, 1)) / (2 * GRAD_VAL)
    Next i
Next j

JACOBI_CENTRAL_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
JACOBI_CENTRAL_FUNC = Err.number
End Function


'// PERFECT

'************************************************************************************
'************************************************************************************
'FUNCTION      : JACOBI_FORWARD_FUNC
'DESCRIPTION   : Compute the Jacobian with the forward step FD formula
'LIBRARY       : OPTIMIZATION
'GROUP         : GRAD_JACOBI
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function JACOBI_FORWARD_FUNC(ByVal FUNC_NAME_STR As String, _
ByRef PARAM_RNG As Variant, _
Optional ByVal epsilon As Double = 10 ^ -5)

Dim i As Long
Dim j As Long
Dim k As Long

Dim NROWS As Long
Dim NSIZE As Long

Dim TEMP_ARR As Variant
Dim XDATA_VECTOR As Variant
Dim YDATA_VECTOR As Variant

Dim XTEMP_VECTOR As Variant
Dim YTEMP_VECTOR As Variant

Dim TEMP_MATRIX As Variant
Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL

PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)
NSIZE = UBound(PARAM_VECTOR, 1)

XDATA_VECTOR = PARAM_VECTOR
ReDim XTEMP_VECTOR(1 To NSIZE, 1 To 1)

TEMP_ARR = Excel.Application.Run(FUNC_NAME_STR, PARAM_VECTOR)
If IsArray(TEMP_ARR) = False Then
    ReDim YDATA_VECTOR(1 To 1, 1 To 1)
    YDATA_VECTOR(1, 1) = TEMP_ARR
Else
    YDATA_VECTOR = TEMP_ARR
    If UBound(YDATA_VECTOR, 1) = 1 Then: YDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(YDATA_VECTOR)
End If

NROWS = UBound(YDATA_VECTOR, 1)

'-------------------------------------------------------------------------------
ReDim TEMP_MATRIX(1 To NROWS, 1 To NSIZE)
'-------------------------------------------------------------------------------
    
For i = 1 To NSIZE
    For j = 1 To NSIZE     'Crea el vector de incr en cada variable
        If j = i Then
            XTEMP_VECTOR(j, 1) = XDATA_VECTOR(j, 1) + epsilon * XDATA_VECTOR(j, 1) + epsilon
        Else
            XTEMP_VECTOR(j, 1) = XDATA_VECTOR(j, 1)
        End If
    Next j
        
    
    TEMP_ARR = Excel.Application.Run(FUNC_NAME_STR, XTEMP_VECTOR)
    If IsArray(TEMP_ARR) = False Then
        ReDim YTEMP_VECTOR(1 To 1, 1 To 1)
        YTEMP_VECTOR(1, 1) = TEMP_ARR
    Else
        YTEMP_VECTOR = TEMP_ARR
        If UBound(YTEMP_VECTOR, 1) = 1 Then: YTEMP_VECTOR = MATRIX_TRANSPOSE_FUNC(YTEMP_VECTOR)
    End If

    For k = 1 To NROWS
        TEMP_MATRIX(k, i) = (YTEMP_VECTOR(k, 1) - YDATA_VECTOR(k, 1)) / (epsilon * XDATA_VECTOR(i, 1) + epsilon)
    Next k
 
 Next i
 
JACOBI_FORWARD_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
JACOBI_FORWARD_FUNC = Err.number
End Function

'// PERFECT


'************************************************************************************
'************************************************************************************
'FUNCTION      : JACOBI_INVERSE_FUNC
'DESCRIPTION   : Compute the inverse of the Jacobian matrix
'LIBRARY       : OPTIMIZATION
'GROUP         : GRAD_JACOBI
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************


Function JACOBI_INVERSE_FUNC(ByVal FUNC_NAME_STR As String, _
ByRef PARAM_RNG As Variant, _
Optional ByVal epsilon As Double = 10 ^ -5)

Dim TEMP_MATRIX As Variant
Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL

PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)

TEMP_MATRIX = JACOBI_FORWARD_FUNC(FUNC_NAME_STR, PARAM_VECTOR, epsilon)
'TEMP_MATRIX = JACOBI_CENTRAL_FUNC(FUNC_NAME_STR, PARAM_VECTOR, epsilon)
TEMP_MATRIX = MATRIX_LU_INVERSE_FUNC(TEMP_MATRIX)
JACOBI_INVERSE_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
JACOBI_INVERSE_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : JACOBI_EVALUATE_MODIFIED_FUNC
'DESCRIPTION   : Evaluate Jacobian of the modified functions
'LIBRARY       : OPTIMIZATION
'GROUP         : GRAD_JACOBI
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function JACOBI_EVALUATE_MODIFIED_FUNC(ByVal FUNC_NAME_STR As String, _
ByRef PARAM_RNG As Variant, _
ByRef ROOTS_RNG As Variant, _
ByVal NO_ROOTS As Long, _
Optional ByVal epsilon As Double = 10 ^ -5)

Dim i As Long
Dim j As Long

Dim NSIZE As Long

Dim GRAD_VAL As Double

Dim TEMP1_VECTOR As Variant
Dim TEMP2_VECTOR As Variant

Dim XDATA_VECTOR As Variant
Dim XTEMP_VECTOR As Variant

Dim TEMP_MATRIX As Variant
Dim PARAM_VECTOR As Variant
Dim ROOTS_MATRIX As Variant

On Error GoTo ERROR_LABEL

ROOTS_MATRIX = ROOTS_RNG
PARAM_VECTOR = PARAM_RNG

If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)
NSIZE = UBound(PARAM_VECTOR, 1)

XDATA_VECTOR = PARAM_VECTOR
ReDim XTEMP_VECTOR(1 To NSIZE, 1 To 1)
ReDim TEMP_MATRIX(1 To NSIZE, 1 To NSIZE)

For j = 1 To NSIZE     'compute jacobian for column forward step
    GRAD_VAL = epsilon * XDATA_VECTOR(j, 1)
    If GRAD_VAL < epsilon Then GRAD_VAL = epsilon
    For i = 1 To NSIZE
        If i = j Then
            XTEMP_VECTOR(i, 1) = XDATA_VECTOR(i, 1) + GRAD_VAL
        Else
            XTEMP_VECTOR(i, 1) = XDATA_VECTOR(i, 1)
        End If
    Next i
    TEMP1_VECTOR = MODIFIED_ROOT_POLES_FUNC(FUNC_NAME_STR, XTEMP_VECTOR, ROOTS_MATRIX, NO_ROOTS, 10 ^ -10)
    GRAD_VAL = epsilon * XDATA_VECTOR(j, 1)
    If GRAD_VAL < epsilon Then GRAD_VAL = epsilon
    For i = 1 To NSIZE
        If i = j Then
            XTEMP_VECTOR(i, 1) = XDATA_VECTOR(i, 1) - GRAD_VAL
        Else
            XTEMP_VECTOR(i, 1) = XDATA_VECTOR(i, 1)
        End If
    Next i
    TEMP2_VECTOR = MODIFIED_ROOT_POLES_FUNC(FUNC_NAME_STR, XTEMP_VECTOR, ROOTS_MATRIX, NO_ROOTS, 10 ^ -10)
    'FD central formula

    For i = 1 To NSIZE
        TEMP_MATRIX(i, j) = (TEMP1_VECTOR(i, 1) - TEMP2_VECTOR(i, 1)) / (2 * GRAD_VAL)
    Next i
Next j

JACOBI_EVALUATE_MODIFIED_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
JACOBI_EVALUATE_MODIFIED_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MODIFIED_ROOT_POLES_FUNC
'DESCRIPTION   : Compute the modified function f(x) with root-poles
'LIBRARY       : OPTIMIZATION
'GROUP         : GRAD_JACOBI
'ID            : 007
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function MODIFIED_ROOT_POLES_FUNC(ByVal FUNC_NAME_STR As String, _
ByRef PARAM_RNG As Variant, _
ByRef ROOTS_RNG As Variant, _
ByVal NO_ROOTS As Long, _
Optional ByVal epsilon As Double = 10 ^ -10)

Dim i As Long
Dim j As Long

Dim NSIZE As Long

Dim TEMP_SUM As Double
Dim MULT_VAL As Double
Dim GRAD_VAL As Double
Dim FACTOR_VAL As Double

Dim ROOTS_MATRIX As Variant
Dim PARAM_VECTOR As Variant
Dim YTEMP_VECTOR As Variant

On Error GoTo ERROR_LABEL

ROOTS_MATRIX = ROOTS_RNG
PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)
NSIZE = UBound(PARAM_VECTOR)  'number of variables

MULT_VAL = 1
For j = 1 To NO_ROOTS
    TEMP_SUM = 0
    For i = 1 To NSIZE
        GRAD_VAL = Abs(ROOTS_MATRIX(j, i) - PARAM_VECTOR(i, 1))
        If GRAD_VAL < 10 ^ 6 Then
            TEMP_SUM = TEMP_SUM + (GRAD_VAL) ^ 2
        Else
            TEMP_SUM = TEMP_SUM + GRAD_VAL
        End If
    Next i
    MULT_VAL = MULT_VAL * TEMP_SUM   '|dx1|*|dx2|*...|dxm|
Next j

FACTOR_VAL = 0.5 * (1 + 1 / (epsilon + MULT_VAL)) 'Durand-Kerner factor

YTEMP_VECTOR = Excel.Application.Run(FUNC_NAME_STR, PARAM_VECTOR)

'rescale functions values
For i = 1 To UBound(YTEMP_VECTOR, 1)
    YTEMP_VECTOR(i, 1) = YTEMP_VECTOR(i, 1) * FACTOR_VAL
Next i

MODIFIED_ROOT_POLES_FUNC = YTEMP_VECTOR

Exit Function
ERROR_LABEL:
MODIFIED_ROOT_POLES_FUNC = Err.number
End Function
