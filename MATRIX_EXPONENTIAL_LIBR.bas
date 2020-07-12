Attribute VB_Name = "MATRIX_EXPONENTIAL_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_PADE_EXPONENTIAL_FUNC

'DESCRIPTION   : This function approximates the exponential of a
'square matrix. This function uses the Padé
'approximation algorithms to approximate the
'infinite summation. It is recommendable especially
'for large matrices.

'LIBRARY       : MATRIX
'GROUP         : EXPONENTIAL
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_PADE_EXPONENTIAL_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal NROWS As Long = 0, _
Optional ByVal epsilon As Double = 10 ^ -15)

'returns the matrix series expansion
'exp(A)= I + A + 1/2*A^2 +1/6*A^3 +...1/n!*A^n + error

Dim k As Long
Dim NCOLUMNS As Long

Dim TEMP_ERR As Double
Dim LOOP_FLAG As Boolean

Dim ATEMP_MATRIX As Variant
Dim BTEMP_MATRIX As Variant
Dim CTEMP_MATRIX As Variant

Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
If UBound(DATA_MATRIX, 1) <> UBound(DATA_MATRIX, 2) Then: GoTo ERROR_LABEL

'series expansion begins
BTEMP_MATRIX = DATA_MATRIX
NCOLUMNS = UBound(DATA_MATRIX, 1)

CTEMP_MATRIX = MATRIX_ELEMENTS_ADD_FUNC(MATRIX_IDENTITY_FUNC(NCOLUMNS), _
            DATA_MATRIX, 1, 1) 'C=I+A

k = 1
GoSub 1983
Do Until LOOP_FLAG
    k = k + 1
    ATEMP_MATRIX = MMULT_FUNC(BTEMP_MATRIX, DATA_MATRIX, 70)
    BTEMP_MATRIX = MATRIX_ELEMENTS_MULT_SCALAR_FUNC(ATEMP_MATRIX, 1 / k)
    CTEMP_MATRIX = MATRIX_ELEMENTS_ADD_FUNC(CTEMP_MATRIX, BTEMP_MATRIX, 1, 1)
    GoSub 1983
Loop

MATRIX_PADE_EXPONENTIAL_FUNC = CTEMP_MATRIX

Exit Function
'---------------------------------------------------------------------------
1983:
'---------------------------------------------------------------------------
If NROWS = 0 Then
    TEMP_ERR = MATRIX_EUCLIDEAN_NORM_FUNC(BTEMP_MATRIX)
    LOOP_FLAG = (TEMP_ERR < epsilon)
Else
    LOOP_FLAG = (k >= NROWS)
End If
'---------------------------------------------------------------------------
Return
'---------------------------------------------------------------------------

ERROR_LABEL:
MATRIX_PADE_EXPONENTIAL_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_POWER_EXPONENTIAL_FUNC

'DESCRIPTION   : This function approximates the exponential of a
'square matrix. This function uses the popular
'power series. For n sufficiently large, the error
'becomes negligible, and the sum approximates the
'exponential matrix function. The parameter nrows
'fixes the max term of the series. If omitted the
'expansion continues until convergence is reached;
'this means that the norm of the nth matrix term
'becomes less than Err = 1e-15..

'When using this function without n,; especially for a
'larger matrix, the evaluation time can be very long.

'LIBRARY       : MATRIX
'GROUP         : EXPONENTIAL
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************


Function MATRIX_POWER_EXPONENTIAL_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal epsilon As Double = 2E-16)

'Scale the matrix by power of 2 so that its norm is < 1/2 .

Dim j As Long
Dim k As Long

Dim NSIZE As Long
Dim NROWS As Long

Dim SWITCH_FLAG As Boolean

Dim TEMP_SCALAR As Double

Dim DATA_MATRIX As Variant
Dim ATEMP_MATRIX As Variant
Dim BTEMP_MATRIX As Variant
Dim CTEMP_MATRIX As Variant
Dim DTEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
NSIZE = MAXIMUM_FUNC(0, Int(Log(MATRIX_EUCLIDEAN_NORM_FUNC(DATA_MATRIX)) / Log(2#)) + 1)

DATA_MATRIX = MATRIX_ELEMENTS_MULT_SCALAR_FUNC(DATA_MATRIX, 0.5 ^ NSIZE)
NROWS = UBound(DATA_MATRIX)

BTEMP_MATRIX = DATA_MATRIX

TEMP_SCALAR = 1 / 2

CTEMP_MATRIX = MATRIX_ELEMENTS_ADD_FUNC(MATRIX_IDENTITY_FUNC(NROWS), _
    MATRIX_ELEMENTS_MULT_SCALAR_FUNC(DATA_MATRIX, TEMP_SCALAR), 1, 1)

DTEMP_MATRIX = MATRIX_ELEMENTS_SUBTRACT_FUNC(MATRIX_IDENTITY_FUNC(NROWS), _
    MATRIX_ELEMENTS_MULT_SCALAR_FUNC(DATA_MATRIX, TEMP_SCALAR), 1, 1)

j = 6

SWITCH_FLAG = True

For k = 2 To j
   TEMP_SCALAR = TEMP_SCALAR * (j - k + 1) / (k * (2 * j - k + 1))
   
   BTEMP_MATRIX = MMULT_FUNC(DATA_MATRIX, BTEMP_MATRIX)
   ATEMP_MATRIX = MATRIX_ELEMENTS_MULT_SCALAR_FUNC(BTEMP_MATRIX, TEMP_SCALAR)
   
   CTEMP_MATRIX = MATRIX_ELEMENTS_ADD_FUNC(CTEMP_MATRIX, ATEMP_MATRIX, 1, 1)
   
   If SWITCH_FLAG Then
         DTEMP_MATRIX = MATRIX_ELEMENTS_ADD_FUNC(DTEMP_MATRIX, ATEMP_MATRIX, 1, 1)
   Else: DTEMP_MATRIX = MATRIX_ELEMENTS_SUBTRACT_FUNC(DTEMP_MATRIX, ATEMP_MATRIX, 1, 1)
   End If
   SWITCH_FLAG = Not (SWITCH_FLAG)
Next k

'CTEMP_MATRIX = MMULT_FUNC(MATRIX_LU_INVERSE_FUNC(DTEMP_MATRIX), CTEMP_MATRIX, 70)
CTEMP_MATRIX = MMULT_FUNC(MATRIX_GS_REDUCTION_PIVOT_FUNC(DTEMP_MATRIX, , epsilon, 1), _
                CTEMP_MATRIX, 70)

'% Undo scaling by repeated squaring
For k = 1 To NSIZE
    CTEMP_MATRIX = MMULT_FUNC(CTEMP_MATRIX, CTEMP_MATRIX, 70)
Next k

MATRIX_POWER_EXPONENTIAL_FUNC = CTEMP_MATRIX
  
Exit Function
ERROR_LABEL:
MATRIX_POWER_EXPONENTIAL_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_EXPONENTIAL_ERROR_FUNC

'DESCRIPTION   : This function returns the truncation n-th term of
'the series of the exponential of a matrix [A]. It is useful to
'estimate the truncation error of the series approximation

'LIBRARY       : MATRIX
'GROUP         : EXPONENTIAL
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_EXPONENTIAL_ERROR_FUNC(ByRef DATA_RNG As Variant, _
ByVal NROWS As Long)

Dim k As Long
Dim NCOLUMNS As Long

Dim TEMP_ERR As Double

Dim DATA_MATRIX As Variant
Dim ATEMP_MATRIX As Variant
Dim BTEMP_MATRIX As Variant
Dim CTEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
If UBound(DATA_MATRIX, 1) <> UBound(DATA_MATRIX, 2) Then: GoTo ERROR_LABEL

BTEMP_MATRIX = DATA_MATRIX
NCOLUMNS = UBound(DATA_MATRIX, 1)
CTEMP_MATRIX = MATRIX_ELEMENTS_ADD_FUNC(MATRIX_IDENTITY_FUNC(NCOLUMNS), _
                DATA_MATRIX, 1, 1)

For k = 2 To NROWS
    ATEMP_MATRIX = MMULT_FUNC(BTEMP_MATRIX, DATA_MATRIX)
    BTEMP_MATRIX = MATRIX_ELEMENTS_MULT_SCALAR_FUNC(ATEMP_MATRIX, 1 / k)
    CTEMP_MATRIX = MATRIX_ELEMENTS_ADD_FUNC(CTEMP_MATRIX, BTEMP_MATRIX, 1, 1)
Next k

TEMP_ERR = MATRIX_EUCLIDEAN_NORM_FUNC(BTEMP_MATRIX)
MATRIX_EXPONENTIAL_ERROR_FUNC = TEMP_ERR
  
Exit Function
ERROR_LABEL:
MATRIX_EXPONENTIAL_ERROR_FUNC = Err.number
End Function
