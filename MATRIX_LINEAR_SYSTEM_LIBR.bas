Attribute VB_Name = "MATRIX_LINEAR_SYSTEM_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_LINEAR_ROWS_COMBINATION_FUNC
'DESCRIPTION   : Linear combination rows
'LIBRARY       : MATRIX
'GROUP         : LINEAR_SYSTEM
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/21/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_LINEAR_ROWS_COMBINATION_FUNC(ByRef DATA_RNG As Variant, _
ByVal ii As Long, _
ByVal jj As Long, _
ByVal SCALAR As Double)
    
Dim j As Long
Dim NCOLUMNS As Long
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
NCOLUMNS = UBound(DATA_MATRIX, 2)
For j = 1 To NCOLUMNS
    DATA_MATRIX(ii, j) = DATA_MATRIX(ii, j) + SCALAR * DATA_MATRIX(jj, j)
Next j

MATRIX_LINEAR_ROWS_COMBINATION_FUNC = DATA_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_LINEAR_ROWS_COMBINATION_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_LINEAR_COLUMNS_COMBINATION_FUNC
'DESCRIPTION   : Linear combination columns
'LIBRARY       : MATRIX
'GROUP         : LINEAR_SYSTEM
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/21/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_LINEAR_COLUMNS_COMBINATION_FUNC(ByRef DATA_RNG As Variant, _
ByVal ii As Long, _
ByVal jj As Long, _
ByVal SCALAR As Double)

Dim i As Long
Dim NROWS As Long
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
NROWS = UBound(DATA_MATRIX, 1)
For i = 1 To NROWS
    DATA_MATRIX(i, ii) = DATA_MATRIX(i, ii) + SCALAR * DATA_MATRIX(i, jj)
Next i

MATRIX_LINEAR_COLUMNS_COMBINATION_FUNC = DATA_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_LINEAR_COLUMNS_COMBINATION_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_LINEAR_TRANSFORMATION_FUNC
'DESCRIPTION   : This function performs the Linear Transformation = Y = Ax + b
'LIBRARY       : MATRIX
'GROUP         : LINEAR_SYSTEM
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/21/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_LINEAR_TRANSFORMATION_FUNC(ByRef MATRIX_RNG As Variant, _
ByRef VECTOR_RNG As Variant, _
Optional ByRef DATA_RNG As Variant, _
Optional ByVal THRESHOLD As Double = 70)

'where:

'MATRIX_RNG (A): is the matrix (n x m) of the transformation
'VECTOR_RNG (x): is the vector of independent variables (m x 1)
'DATA_RNG (b): is the known vector (n x 1)

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_MATRIX As Variant

Dim XDATA_VECTOR As Variant
Dim YDATA_VECTOR As Variant
'is the vector of dependent variables (n x 1)

Dim DATA1_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA1_MATRIX = MATRIX_RNG
XDATA_VECTOR = VECTOR_RNG
YDATA_VECTOR = MMULT_FUNC(DATA1_MATRIX, XDATA_VECTOR, THRESHOLD)

NROWS = UBound(YDATA_VECTOR, 1)
NCOLUMNS = UBound(YDATA_VECTOR, 2)

If IsArray(DATA_RNG) = True Then
    TEMP_MATRIX = DATA_RNG
    For i = 1 To NROWS
        For j = 1 To NCOLUMNS
           YDATA_VECTOR(i, j) = YDATA_VECTOR(i, j) + TEMP_MATRIX(i, j)
        Next j
    Next i
End If

MATRIX_LINEAR_TRANSFORMATION_FUNC = YDATA_VECTOR

Exit Function
ERROR_LABEL:
MATRIX_LINEAR_TRANSFORMATION_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_LEAST_SQUARE_LINEAR_SYSTEM_FUNC
'DESCRIPTION   : least square linear regression matrix
'LIBRARY       : MATRIX
'GROUP         : LINEAR_SYSTEM
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_LEAST_SQUARE_LINEAR_SYSTEM_FUNC(ByRef XDATA_RNG As Variant, _
ByRef YDATA_RNG As Variant, _
Optional ByVal VERSION As Integer = 0)

Dim i As Long
Dim j As Long
Dim k As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim YDATA_VECTOR As Variant
Dim XDATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

XDATA_MATRIX = XDATA_RNG
NROWS = UBound(XDATA_MATRIX, 1)
NCOLUMNS = UBound(XDATA_MATRIX, 2)

YDATA_VECTOR = YDATA_RNG
If UBound(YDATA_VECTOR, 1) = 1 Then: YDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(YDATA_VECTOR)
If NROWS <> UBound(YDATA_VECTOR, 1) Then: GoTo ERROR_LABEL

'----------------------------------------------------------------------------------
Select Case VERSION
'----------------------------------------------------------------------------------
Case 0 'useful for solving least square linear regression x*c = y
'----------------------------------------------------------------------------------
    ReDim TEMP_MATRIX(1 To NCOLUMNS, 1 To NCOLUMNS + 1)
    
    For i = 1 To NCOLUMNS
        For j = i To NCOLUMNS
            For k = 1 To NROWS
                TEMP_MATRIX(i, j) = TEMP_MATRIX(i, j) + XDATA_MATRIX(k, i) * XDATA_MATRIX(k, j)
            Next k
            If i <> j Then TEMP_MATRIX(j, i) = TEMP_MATRIX(i, j)
        Next j
    Next i
    
    For i = 1 To NCOLUMNS
        For k = 1 To NROWS
            TEMP_MATRIX(i, NCOLUMNS + 1) = TEMP_MATRIX(i, NCOLUMNS + 1) + YDATA_VECTOR(k, 1) * XDATA_MATRIX(k, i)
        Next k
    Next i
    
'----------------------------------------------------------------------------------
Case Else 'useful for solving linear system a*x = b
'----------------------------------------------------------------------------------
    ReDim TEMP_MATRIX(1 To NROWS, 1 To NROWS + 1) 'build the system matrix
    For i = 1 To NROWS
        For j = 1 To NROWS
            TEMP_MATRIX(i, j) = XDATA_MATRIX(i, j)
        Next j
    Next i
    For i = 1 To NROWS
        TEMP_MATRIX(i, NROWS + 1) = YDATA_VECTOR(i, 1)
    Next i
'----------------------------------------------------------------------------------
End Select
'----------------------------------------------------------------------------------

MATRIX_LEAST_SQUARE_LINEAR_SYSTEM_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_LEAST_SQUARE_LINEAR_SYSTEM_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_TRIANGULAR_LINEAR_SYSTEM_FUNC
'DESCRIPTION   : Solving for linear systems

'LIBRARY       : MATRIX
'GROUP         : LINEAR SYSTEM
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_TRIANGULAR_LINEAR_SYSTEM_FUNC(ByRef DATA_RNG As Variant)

Dim i As Long
Dim j As Long
Dim k As Long

Dim ii As Long
Dim jj As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim MAX_VAL As Double
Dim DET_VAL As Double
Dim TEMP_VAL As Double

Dim TEMP_ARR() As Double
Dim COEF_VECTOR() As Double

Dim TEMP1_MATRIX() As Double
Dim TEMP2_MATRIX() As Double

Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

ReDim TEMP1_MATRIX(1 To NROWS, 1 To NCOLUMNS)
For j = 1 To NCOLUMNS
    For i = 1 To NROWS
        TEMP1_MATRIX(i, j) = DATA_MATRIX(i, j)
    Next i
Next j

ReDim COEF_VECTOR(1 To NROWS, 1 To 1)
ReDim TEMP2_MATRIX(1 To NROWS, 1 To NROWS)

For k = 1 To NROWS
    ReDim TEMP_ARR(k To NROWS)
    For ii = k To NROWS
        TEMP_ARR(ii) = Abs(TEMP1_MATRIX(ii, k))
        'For jj = k + 1 To NROWS + 1
        For jj = k + 1 To NROWS
            If Abs(TEMP1_MATRIX(ii, jj)) > TEMP_ARR(ii) Then
                TEMP_ARR(ii) = Abs(TEMP1_MATRIX(ii, jj))
            End If
        Next jj
    Next ii

    For ii = k To NROWS
        If TEMP_ARR(ii) <> 0 Then
            TEMP_ARR(ii) = Abs(TEMP1_MATRIX(ii, k)) / TEMP_ARR(ii)
        End If
    Next ii
    
    MAX_VAL = TEMP_ARR(k)
    If MAX_VAL = 0 Then GoTo 1983
    For ii = k + 1 To NROWS
        If MAX_VAL < TEMP_ARR(ii) Then
            For jj = k To NROWS + 1
                TEMP_VAL = TEMP1_MATRIX(k, jj)
                TEMP1_MATRIX(k, jj) = TEMP1_MATRIX(ii, jj)
                TEMP1_MATRIX(ii, jj) = TEMP_VAL
            Next jj
            MAX_VAL = TEMP_ARR(ii)
        End If
    Next ii
1983:
    For i = k + 1 To NROWS
        If TEMP1_MATRIX(k, k) = 0 Then GoTo 1984
        TEMP2_MATRIX(i, k) = TEMP1_MATRIX(i, k) / TEMP1_MATRIX(k, k)
        TEMP1_MATRIX(i, k) = 0
        For j = k + 1 To NROWS + 1
            TEMP1_MATRIX(i, j) = TEMP1_MATRIX(i, j) - TEMP2_MATRIX(i, k) * TEMP1_MATRIX(k, j)
        Next j
    Next i
Next k
1984:

' Aseguramos que el sistema tiene solucion
DET_VAL = 1
For k = 1 To NROWS
    DET_VAL = TEMP1_MATRIX(k, k) * DET_VAL
Next k
If DET_VAL <> 0 Then
    For k = NROWS To 1 Step -1
        COEF_VECTOR(k, 1) = (TEMP1_MATRIX(k, NROWS + 1) / TEMP1_MATRIX(k, k))
    Next k
    For k = NROWS - 1 To 1 Step -1
        For i = k + 1 To NROWS
            COEF_VECTOR(k, 1) = (COEF_VECTOR(k, 1) - (TEMP1_MATRIX(k, i) / TEMP1_MATRIX(k, k)) * COEF_VECTOR(i, 1))
        Next i
    Next k
End If

MATRIX_TRIANGULAR_LINEAR_SYSTEM_FUNC = Array(COEF_VECTOR, DET_VAL)

Exit Function
ERROR_LABEL:
MATRIX_TRIANGULAR_LINEAR_SYSTEM_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_JACOBI_LINEAR_SYSTEM_FUNC

'DESCRIPTION   : This function performs the iterative Jacobi algorithm for linear
'system solving and has been developed for its didactic scope in order to study
'the convergence of iterative process. The function returns the vector after Nmax
'steps; if the matrix is convergent, this vector is closer to the exact solution.

'LIBRARY       : MATRIX
'GROUP         : LINEAR SYSTEM
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_JACOBI_LINEAR_SYSTEM_FUNC(ByRef DATA1_RNG As Variant, _
ByRef DATA2_RNG As Variant, _
Optional ByRef YDATA_RNG As Variant, _
Optional ByVal nLOOPS As Long = 200)

'DATA1_RNG:  is the system matrix
'DATA2_RNG: is the constant term (n TEMP2_MATRIX 1) vector
'YDATA_RNG: is the (nx1) vector of the starting approximate solution
'nLOOPS:       is the max step allowed (default = 200)

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_VAL As Double

Dim YDATA_VECTOR As Variant

Dim DATA1_MATRIX As Variant
Dim DATA2_MATRIX As Variant

Dim TEMP1_MATRIX As Variant
Dim TEMP2_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA1_MATRIX = DATA1_RNG
DATA2_MATRIX = DATA2_RNG
NROWS = UBound(DATA1_MATRIX, 1)
NCOLUMNS = UBound(DATA1_MATRIX, 2)
ReDim TEMP2_MATRIX(1 To UBound(DATA2_MATRIX, 1), 1 To UBound(DATA2_MATRIX, 2))

If NCOLUMNS <> NROWS Or NROWS <> UBound(DATA2_MATRIX, 1) Or UBound(DATA2_MATRIX, 2) <> 1 Then: GoTo ERROR_LABEL

If IsArray(YDATA_RNG) = False Then 'load starting vector
    For i = 1 To UBound(DATA2_MATRIX, 1)
        TEMP2_MATRIX(i, 1) = 0
    Next i
Else
    YDATA_VECTOR = YDATA_RNG
    If UBound(YDATA_VECTOR, 1) = 1 Then
        YDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(YDATA_VECTOR)
    End If
    For i = 1 To UBound(DATA2_MATRIX, 1)
        TEMP2_MATRIX(i, 1) = YDATA_VECTOR(i, 1)
    Next i
End If

NROWS = UBound(DATA1_MATRIX, 1)
NCOLUMNS = UBound(DATA2_MATRIX, 2)

ReDim TEMP1_MATRIX(1 To NROWS, 1 To NCOLUMNS)

For h = 1 To nLOOPS
    For j = 1 To NCOLUMNS
        For i = 1 To NROWS
            TEMP1_MATRIX(i, j) = TEMP2_MATRIX(i, j)
        Next i
    Next j

    For k = 1 To NCOLUMNS
        For i = 1 To NROWS
            TEMP_VAL = DATA2_MATRIX(i, k)
            For j = 1 To NROWS
                If i <> j Then: TEMP_VAL = TEMP_VAL - _
                    DATA1_MATRIX(i, j) * TEMP1_MATRIX(j, k)
            Next j
            TEMP2_MATRIX(i, k) = TEMP_VAL / DATA1_MATRIX(i, i)
        Next i
    Next k
Next h

MATRIX_JACOBI_LINEAR_SYSTEM_FUNC = TEMP2_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_JACOBI_LINEAR_SYSTEM_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_SEIDEL_LINEAR_SYSTEM_FUNC

'DESCRIPTION   : This function performs the iterative Gauss-Seidel algorithm
'for linear system solving and has been developed for its didactic scope in
'order to study the convergence of iterative processes.

'The function returns the vector after nLOOPS steps; if the matrix is
'convergent, this vector is closer to the exact solution.


'LIBRARY       : MATRIX
'GROUP         : LINEAR SYSTEM
'ID            : 007
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_SEIDEL_LINEAR_SYSTEM_FUNC(ByRef DATA1_RNG As Variant, _
ByRef DATA2_RNG As Variant, _
Optional ByRef YDATA_RNG As Variant, _
Optional ByVal RELAX_VAL As Double = 1, _
Optional ByVal nLOOPS As Long = 200)

'DATA1_RNG:  is the system matrix
'DATA2_RNG: is the constant term (n TEMP2_MATRIX 1) vector
'YDATA_RNG: is the (nx1) vector of the starting approximate solution
'nLOOPS:       is the max step allowed (default = 200)

'Precision increases with an increasing number of steps (of course,
'for a convergent matrix). Usually, the convergence speed is quite
'low, but it can be greatly accelerate by the Aitken extrapolation
'formula, also called "square delta extrapolation", or by tuning the
'relaxation factor RELAX_VAL

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_VAL As Double

Dim DATA1_MATRIX As Variant
Dim DATA2_MATRIX As Variant

Dim TEMP2_MATRIX As Variant
Dim YDATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA1_MATRIX = DATA1_RNG
DATA2_MATRIX = DATA2_RNG

NROWS = UBound(DATA1_MATRIX, 1)
NCOLUMNS = UBound(DATA1_MATRIX, 2)

ReDim TEMP2_MATRIX(1 To UBound(DATA2_MATRIX, 1), 1 To 1)

'If NCOLUMNS <> NROWS Or NROWS <> UBound(DATA2_MATRIX, 1) _
    Or UBound(DATA2_MATRIX, 2) <> 1 Then: GoTo ERROR_LABEL

If IsArray(YDATA_VECTOR) = False Then 'load starting vector
    For i = 1 To UBound(DATA2_MATRIX, 1)
        TEMP2_MATRIX(i, 1) = 0
    Next i
Else
    YDATA_VECTOR = YDATA_RNG
    If UBound(YDATA_VECTOR, 1) = 1 Then
        YDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(YDATA_VECTOR)
    End If
    For i = 1 To UBound(DATA2_MATRIX, 1)
        TEMP2_MATRIX(i, 1) = YDATA_VECTOR(i, 1)
    Next i
End If

NROWS = UBound(DATA1_MATRIX, 1)
NCOLUMNS = UBound(DATA2_MATRIX, 2)

'-----------------------------I/h Diagonal Matrix--------------------------------
For i = 1 To NROWS
    TEMP_VAL = DATA1_MATRIX(i, i)
    For k = 1 To NCOLUMNS
        DATA2_MATRIX(i, k) = DATA2_MATRIX(i, k) / TEMP_VAL
    Next k
    For j = 1 To NROWS
        DATA1_MATRIX(i, j) = DATA1_MATRIX(i, j) / TEMP_VAL
    Next j
Next i

For h = 1 To nLOOPS
    For k = 1 To NCOLUMNS
        For i = 1 To NROWS
            TEMP_VAL = 0
            For j = 1 To NROWS
                TEMP_VAL = TEMP_VAL + DATA1_MATRIX(i, j) * TEMP2_MATRIX(j, k)
            Next j
            TEMP_VAL = DATA2_MATRIX(i, k) - TEMP_VAL
            TEMP2_MATRIX(i, k) = TEMP2_MATRIX(i, k) + RELAX_VAL * TEMP_VAL
        Next i
    Next k
Next h
'--------------------------------------------------------------------------------


MATRIX_SEIDEL_LINEAR_SYSTEM_FUNC = TEMP2_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_SEIDEL_LINEAR_SYSTEM_FUNC = Err.number
End Function

