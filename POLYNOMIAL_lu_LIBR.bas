Attribute VB_Name = "POLYNOMIAL_lu_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : POLYNOMIAL_REGRESSION_FUNC
'DESCRIPTION   : Least-squares regression with polynomials: By solving
'the linear system of equations using LU Factorization
'LIBRARY       : POLYNOMIAL
'GROUP         : LU
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function POLYNOMIAL_REGRESSION_FUNC(ByRef XDATA_RNG As Variant, _
ByRef YDATA_RNG As Variant, _
ByRef NDEG As Integer, _
Optional ByVal NROWS As Long = 0)
  
Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim ii As Long
Dim jj As Long

Dim XDATA_VECTOR As Variant
Dim YDATA_VECTOR As Variant

Dim TEMP_SUM As Double
Dim TEMP_MATRIX As Variant
Dim TEMP_VECTOR As Variant
Dim COEF_VECTOR As Variant

Dim tolerance As Double

On Error GoTo ERROR_LABEL

'-----------------------------------------------------------------------------
XDATA_VECTOR = XDATA_RNG
If UBound(XDATA_VECTOR, 1) = 1 Then
    XDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(XDATA_VECTOR)
End If
YDATA_VECTOR = YDATA_RNG
If UBound(YDATA_VECTOR, 1) = 1 Then
    YDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(YDATA_VECTOR)
End If
If UBound(YDATA_VECTOR, 1) <> UBound(YDATA_VECTOR, 1) Then: GoTo ERROR_LABEL

'------------------> Degree of polynomial cannot exceed NROWS - 1

tolerance = 0.000000000000001

If NROWS = 0 Then: NROWS = UBound(XDATA_VECTOR, 1)
If NDEG > (NROWS - 1) Then: NDEG = NROWS - 1
If NDEG > 7 Then: NDEG = 7

'-----------------------------------------------------------------------------
 
' Definition and initialization arrays
ReDim TEMP_MATRIX(1 To NDEG + 1, 1 To NDEG + 2)
ReDim COEF_VECTOR(1 To NDEG + 1, 1 To 1)  ' array of the coefficients of
ReDim TEMP_VECTOR(1 To NROWS)

For i = 1 To NROWS: TEMP_VECTOR(i) = 1: Next i
'-------------First Pass: Compute 1st column and N+1 st column of TEMP_MATRIX

For i = 1 To NDEG + 1
    TEMP_MATRIX(i, 1) = 0
    TEMP_MATRIX(i, NDEG + 2) = 0
    For j = 1 To NROWS
        TEMP_MATRIX(i, 1) = TEMP_MATRIX(i, 1) + TEMP_VECTOR(j)
        TEMP_MATRIX(i, NDEG + 2) = TEMP_MATRIX(i, NDEG + 2) + YDATA_VECTOR(j, 1) * _
        TEMP_VECTOR(j)
        TEMP_VECTOR(j) = TEMP_VECTOR(j) * XDATA_VECTOR(j, 1)
    Next j
Next i

'----------------Second Pass: Compute the last row of TEMP_MATRIX

For i = 2 To NDEG + 1
    TEMP_MATRIX(NDEG, i) = 0
    For j = 1 To NROWS
        TEMP_MATRIX(NDEG + 1, i) = TEMP_MATRIX(NDEG + 1, i) + TEMP_VECTOR(j)
        TEMP_VECTOR(j) = TEMP_VECTOR(j) * XDATA_VECTOR(j, 1)
    Next j
Next i

'--------------------Third Pass: Fill rest of matrix
For j = 2 To NDEG + 1
    For i = 1 To NDEG
        TEMP_MATRIX(i, j) = TEMP_MATRIX(i + 1, j - 1)
    Next i
Next j

'----------Forth Pass: forms the LU equivalent of the square coefficient
'matrix A.

For i = 1 To NDEG + 1
    For j = 2 To NDEG + 1
        TEMP_SUM = 0
        If (j <= i) Then
            l = j - 1
            For k = 1 To l
                TEMP_SUM = TEMP_SUM + TEMP_MATRIX(i, k) * TEMP_MATRIX(k, j)
            Next k
            TEMP_MATRIX(i, j) = TEMP_MATRIX(i, j) - TEMP_SUM
        Else
            l = i - 1
            If (l <> 0) Then
                For k = 1 To l
                    TEMP_SUM = TEMP_SUM + TEMP_MATRIX(i, k) * TEMP_MATRIX(k, j)
                Next k
            End If
            If (Abs(TEMP_MATRIX(i, i)) < tolerance) Then: GoTo ERROR_LABEL
            TEMP_MATRIX(i, j) = (TEMP_MATRIX(i, j) - TEMP_SUM) / TEMP_MATRIX(i, i)
        End If
    Next j
Next i

'---------Fifth Pass: Coeficient matrix needs to be LU equivalent of the
'---------square coefficient matrix.

For j = 1 To NDEG + 1
    COEF_VECTOR(j, 1) = TEMP_MATRIX(j, NDEG + 2)
Next j

'----------Sixth Pass: Finds the solution to a set of N linear equations
'----------that correspond to the right-hand side vector.

COEF_VECTOR(1, 1) = COEF_VECTOR(1, 1) / TEMP_MATRIX(1, 1)
For i = 2 To NDEG + 1
    l = i - 1
    TEMP_SUM = 0
    For k = 1 To l
        TEMP_SUM = TEMP_SUM + TEMP_MATRIX(i, k) * COEF_VECTOR(k, 1)
    Next k
    COEF_VECTOR(i, 1) = (COEF_VECTOR(i, 1) - TEMP_SUM) / TEMP_MATRIX(i, i)
Next i

For j = 2 To NDEG + 1   ' now back substitution
    jj = NDEG + 1 - j + 2
    ii = NDEG + 1 - j + 1
    TEMP_SUM = 0
    For k = jj To NDEG + 1
        TEMP_SUM = TEMP_SUM + TEMP_MATRIX(ii, k) * COEF_VECTOR(k, 1)
    Next k
    COEF_VECTOR(ii, 1) = COEF_VECTOR(ii, 1) - TEMP_SUM
Next j

POLYNOMIAL_REGRESSION_FUNC = COEF_VECTOR

Exit Function
ERROR_LABEL:
POLYNOMIAL_REGRESSION_FUNC = Err.number
End Function
