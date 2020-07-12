Attribute VB_Name = "OPTIM_NLR_LMA_LIBR"


'// PERFECT

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : LEVENBERG_MARQUARDT_OPTIMIZATION1_FUNC

'DESCRIPTION   : This is a popular alternative to the Gauss-Newton
'method of finding the minimum of a function that is a sum of squares
'of nonlinear functions. This algorithm was found to be an efficient,
'fast and robust method which also has a good global convergence
'property.

'This function has several routines for performing least squares fitting of
'nonlinear functions directly on the worksheet with the Levenberg-Marquardt
'algorithm3. It uses the derivatives information (if available) or approximates
'them internally by the finite difference method. It needs also the function
'definition cell (=f(x, p1, p2,...), the parameter starting values
'(p1, p2,...), and of course the dataset (xi, yi).

'The Levenberg-Marquardt method uses a search direction that is a cross between
'the Gauss-Newton direction and the steepest descent. It is an heuristic method
'that works extremely well in practice. It has become a virtual standard for
'optimization of medium sized nonlinear models.

' REFERENCES:
' Numerical Recipies in Fortran77; W.H. Press, et al.; Cambridge U. Press
' Pages 683 - 685
' Metodos numericos con Matlab; J. M Mathewss et al.; Prentice Hall

'LIBRARY       : OPTIMIZATION
'GROUP         : NLR_LMA
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function LEVENBERG_MARQUARDT_OPTIMIZATION1_FUNC(ByRef XDATA_RNG As Variant, _
ByRef YDATA_RNG As Variant, _
ByRef PARAM_RNG As Variant, _
ByVal FUNC_NAME_STR As String, _
Optional ByVal GRAD_STR_NAME As String = "", _
Optional ByVal OUTPUT As Integer = 0, _
Optional ByVal nLOOPS As Variant = 1000, _
Optional ByVal tolerance As Double = 10 ^ -15, _
Optional ByVal epsilon As Double = 10 ^ -10)

Dim i As Long
Dim h As Long
    
Dim NROWS As Long
Dim NCOLUMNS As Long
Dim NSIZE As Long

Dim nSTEPS As Long
Dim COUNTER As Long

Dim TEMP_SUM As Double
Dim TEMP_FACT As Double          ' step factor
Dim TEMP_MULT As Double         ' Valor delta de algoritmo

Dim ATEMP_DELTA As Double            ' reduction/amplification quadratic error
Dim BTEMP_DELTA As Double        ' relative increment

Dim ATEMP_VALUE As Double          ' quadratic error previous step
Dim BTEMP_VALUE As Double

Dim XDATA_MATRIX As Variant
Dim PARAM_VECTOR As Variant
Dim YDATA_VECTOR As Variant

Dim DELTA_VECTOR As Variant
Dim BETA_VECTOR As Variant
Dim FIT_VECTOR As Variant

Dim COEF_MATRIX As Variant

Dim MID_BOUND As Double        ' increment previous step
Dim LOWER_BOUND As Double

On Error GoTo ERROR_LABEL

XDATA_MATRIX = XDATA_RNG
NROWS = UBound(XDATA_MATRIX, 1)
NCOLUMNS = UBound(XDATA_MATRIX, 2)

YDATA_VECTOR = YDATA_RNG
If UBound(YDATA_VECTOR, 1) = 1 Then: YDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(YDATA_VECTOR)
If UBound(YDATA_VECTOR, 1) <> NROWS Then: GoTo ERROR_LABEL

PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)
NSIZE = UBound(PARAM_VECTOR, 1)

h = 0
LOWER_BOUND = 10 ^ -3
   
TEMP_MULT = LOWER_BOUND
   
COUNTER = 0
TEMP_FACT = 10
nSTEPS = nLOOPS / 4

FIT_VECTOR = LEVENBERG_MARQUARDT_OBJECTIVE_FUNC(XDATA_MATRIX, _
                 PARAM_VECTOR, FUNC_NAME_STR)
        
TEMP_SUM = 0
For i = 1 To NROWS
    TEMP_SUM = TEMP_SUM + (YDATA_VECTOR(i, 1) - FIT_VECTOR(i, 1)) ^ 2
Next i
    
ATEMP_VALUE = TEMP_SUM

' Rutina para calcular los coeficintes de la matriz asociada al
' ajuste por minimos cuadrados para un modelo mediante el algoritmo
' de Lebenberg-Marquart

Do While h < nSTEPS And COUNTER < nLOOPS

    FIT_VECTOR = LEVENBERG_MARQUARDT_OBJECTIVE_FUNC(XDATA_MATRIX, _
                  PARAM_VECTOR, FUNC_NAME_STR)

    If GRAD_STR_NAME = "" Then
        COEF_MATRIX = LEVENBERG_MARQUARDT_GRAD_FD_APPROX_FUNC(XDATA_MATRIX, PARAM_VECTOR, _
                  FUNC_NAME_STR, epsilon)
    Else
        COEF_MATRIX = LEVENBERG_MARQUARDT_GRAD_FUNC(XDATA_MATRIX, PARAM_VECTOR, _
                  GRAD_STR_NAME)
    End If

    BETA_VECTOR = MMULT_FUNC(MATRIX_TRANSPOSE_FUNC(COEF_MATRIX), _
                  MATRIX_ELEMENTS_SUBTRACT_FUNC(YDATA_VECTOR, _
                  FIT_VECTOR, 1, 1), 70)
                  
    COEF_MATRIX = MMULT_FUNC(MATRIX_TRANSPOSE_FUNC(COEF_MATRIX), COEF_MATRIX, 70)

    For i = 1 To NSIZE
        COEF_MATRIX(i, i) = COEF_MATRIX(i, i) * (1 + TEMP_MULT)
    Next i
    
    
    DELTA_VECTOR = MATRIX_LU_LINEAR_SYSTEM_FUNC(COEF_MATRIX, BETA_VECTOR)

    For i = 1 To NSIZE
        DELTA_VECTOR(i, 1) = PARAM_VECTOR(i, 1) + _
                             DELTA_VECTOR(i, 1)
    Next i
            
    FIT_VECTOR = LEVENBERG_MARQUARDT_OBJECTIVE_FUNC(XDATA_MATRIX, _
                  DELTA_VECTOR, FUNC_NAME_STR)
    
    TEMP_SUM = 0
    For i = 1 To NROWS
        TEMP_SUM = TEMP_SUM + (YDATA_VECTOR(i, 1) - FIT_VECTOR(i, 1)) ^ 2
    Next i
    
    BTEMP_VALUE = TEMP_SUM
    
    If BTEMP_VALUE > tolerance Then ATEMP_DELTA = BTEMP_VALUE / ATEMP_VALUE
    
    BTEMP_DELTA = (BTEMP_VALUE - ATEMP_VALUE)
    If BTEMP_VALUE > tolerance Then BTEMP_DELTA = BTEMP_DELTA / BTEMP_VALUE
    
    If ATEMP_DELTA > 10 Then
        TEMP_MULT = TEMP_MULT * TEMP_FACT
    Else
        If Abs(BTEMP_DELTA) < LOWER_BOUND Then h = h + 1
        TEMP_MULT = TEMP_MULT / TEMP_FACT
        For i = 1 To NSIZE
            PARAM_VECTOR(i, 1) = DELTA_VECTOR(i, 1) 'update new point
        Next i
        ATEMP_VALUE = BTEMP_VALUE
    End If
    
    If COUNTER Mod 4 = 0 Then 'check oscillation
        If MID_BOUND = TEMP_MULT Then
              TEMP_FACT = 2  'relaxed step
        Else: TEMP_FACT = 10 'fast step
        End If
        MID_BOUND = TEMP_MULT
    End If
    
    COUNTER = COUNTER + 1
Loop
    
            
'----------------------------------------------------------------------------------------
Select Case OUTPUT
'----------------------------------------------------------------------------------------
Case 0
    LEVENBERG_MARQUARDT_OPTIMIZATION1_FUNC = PARAM_VECTOR 'Final Parameters
'----------------------------------------------------------------------------------------
Case 1
    
    LEVENBERG_MARQUARDT_OPTIMIZATION1_FUNC = LEVENBERG_MARQUARDT_OBJECTIVE_FUNC(XDATA_MATRIX, _
         PARAM_VECTOR, FUNC_NAME_STR)

'----------------------------------------------------------------------------------------
Case 2
    LEVENBERG_MARQUARDT_OPTIMIZATION1_FUNC = LEVENBERG_MARQUARDT_GRAD_FUNC(XDATA_MATRIX, _
              PARAM_VECTOR, GRAD_STR_NAME)
'----------------------------------------------------------------------------------------
Case 3
    
    LEVENBERG_MARQUARDT_OPTIMIZATION1_FUNC = LEVENBERG_MARQUARDT_GRAD_FD_VALID_FUNC(XDATA_RNG, _
              PARAM_VECTOR, FUNC_NAME_STR, GRAD_STR_NAME, _
              tolerance, epsilon)
'----------------------------------------------------------------------------------------
Case 4
    LEVENBERG_MARQUARDT_OPTIMIZATION1_FUNC = LEVENBERG_MARQUARDT_SIGMA_FUNC(XDATA_RNG, _
                              YDATA_RNG, PARAM_VECTOR, FUNC_NAME_STR, _
                              GRAD_STR_NAME, 0, epsilon)

'----------------------------------------------------------------------------------------
Case Else
    LEVENBERG_MARQUARDT_OPTIMIZATION1_FUNC = COUNTER
'----------------------------------------------------------------------------------------
End Select
'----------------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
LEVENBERG_MARQUARDT_OPTIMIZATION1_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : LEVENBERG_MARQUARDT_OPTIMIZATION2_FUNC

'DESCRIPTION   : This is a popular alternative to the Gauss-Newton
'method of finding the minimum of a function that is a sum of squares
'of nonlinear functions. This algorithm was found to be an efficient,
'fast and robust method which also has a good global convergence
'property.

'This function has several routines for performing least squares fitting of
'nonlinear functions directly on the worksheet with the Levenberg-Marquardt
'algorithm3. It uses the derivatives information (if available) or approximates
'them internally by the finite difference method. It needs also the function
'definition cell (=f(x, p1, p2,...), the parameter starting values
'(p1, p2,...), and of course the dataset (xi, yi).

'The Levenberg-Marquardt method uses a search direction that is a cross between
'the Gauss-Newton direction and the steepest descent. It is an heuristic method
'that works extremely well in practice. It has become a virtual standard for
'optimization of medium sized nonlinear models.

' REFERENCES:
' Numerical Recipies in Fortran77; W.H. Press, et al.; Cambridge U. Press
' Pages 683 - 685
' Metodos numericos con Matlab; J. M Mathewss et al.; Prentice Hall

'LIBRARY       : OPTIMIZATION
'GROUP         : NLR_LMA
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function LEVENBERG_MARQUARDT_OPTIMIZATION2_FUNC(ByRef XDATA_RNG As Variant, _
ByRef YDATA_RNG As Variant, _
ByRef PARAM_RNG As Variant, _
ByVal FUNC_NAME_STR As String, _
Optional ByVal GRAD_STR_NAME As String = "", _
Optional ByVal OUTPUT As Integer = 0, _
Optional ByVal nLOOPS As Variant = 1000, _
Optional ByVal tolerance As Double = 10 ^ -15, _
Optional ByVal epsilon As Double = 10 ^ -10)

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long

Dim NROWS As Long
Dim NCOLUMNS As Long
Dim NSIZE As Long

Dim nSTEPS As Long
Dim COUNTER As Long

Dim TEMP_SUM As Double
Dim TEMP_FACT As Double ' step factor
Dim TEMP_MULT As Double

Dim ATEMP_VALUE As Double
Dim BTEMP_VALUE As Double ' quadratic error previous step

Dim ATEMP_DELTA As Double ' reduction/amplification quadratic error
Dim BTEMP_DELTA As Double ' relative increment

Dim TEMP_ARR As Variant

Dim COEF_MATRIX As Variant
Dim DELTA_VECTOR As Variant
Dim GRAD_MATRIX As Variant
Dim FIT_VECTOR As Variant

Dim XDATA_MATRIX As Variant
Dim YDATA_VECTOR As Variant
Dim PARAM_VECTOR As Variant

Dim MID_BOUND As Double ' increment previous step
Dim LOWER_BOUND As Double


On Error GoTo ERROR_LABEL

XDATA_MATRIX = XDATA_RNG
NROWS = UBound(XDATA_MATRIX, 1)
NCOLUMNS = UBound(XDATA_MATRIX, 2)

YDATA_VECTOR = YDATA_RNG
If UBound(YDATA_VECTOR, 1) = 1 Then: YDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(YDATA_VECTOR)
If UBound(YDATA_VECTOR, 1) <> NROWS Then: GoTo ERROR_LABEL

PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)
NSIZE = UBound(PARAM_VECTOR, 1)

h = 0
LOWER_BOUND = 10 ^ -3

TEMP_MULT = LOWER_BOUND  '0.1%
COUNTER = 0
TEMP_FACT = 10
nSTEPS = nLOOPS / 4
    
FIT_VECTOR = LEVENBERG_MARQUARDT_OBJECTIVE_FUNC(XDATA_MATRIX, _
             PARAM_VECTOR, FUNC_NAME_STR)

    'take all function values
TEMP_SUM = 0
For i = 1 To NROWS
    TEMP_SUM = TEMP_SUM + _
                    (YDATA_VECTOR(i, 1) - FIT_VECTOR(i, 1)) ^ 2
Next i
    
ATEMP_VALUE = TEMP_SUM

Do While h < nSTEPS And COUNTER < nLOOPS
        
    ReDim COEF_MATRIX(1 To NSIZE, 1 To NSIZE + 1)
   ' Rutina para calcular los coeficintes de la matriz asociada
   ' al ajuste por minimos cuadrados para un modelo
   ' mediante el algoritmo de Lebenberg-Marquart
            
    FIT_VECTOR = LEVENBERG_MARQUARDT_OBJECTIVE_FUNC(XDATA_MATRIX, _
                      PARAM_VECTOR, FUNC_NAME_STR)
                      
                  
    If GRAD_STR_NAME = "" Then
        GRAD_MATRIX = LEVENBERG_MARQUARDT_GRAD_FD_APPROX_FUNC(XDATA_MATRIX, PARAM_VECTOR, _
                      FUNC_NAME_STR, epsilon)
    Else
        GRAD_MATRIX = LEVENBERG_MARQUARDT_GRAD_FUNC(XDATA_MATRIX, PARAM_VECTOR, _
                      GRAD_STR_NAME)
    End If

    For i = 1 To NROWS
        For k = 1 To NSIZE
            COEF_MATRIX(k, NSIZE + 1) = _
                (YDATA_VECTOR(i, 1) - FIT_VECTOR(i, 1)) * _
                 GRAD_MATRIX(i, k) + COEF_MATRIX(k, NSIZE + 1)
                        
            For j = k To NSIZE
                    COEF_MATRIX(k, j) = GRAD_MATRIX(i, k) * GRAD_MATRIX(i, j) _
                                    + COEF_MATRIX(k, j)
            Next j
        Next k
    Next i
    For k = 2 To NSIZE
        For j = 1 To k - 1
            COEF_MATRIX(k, j) = COEF_MATRIX(j, k)
        Next j
    Next k
    
    For i = 1 To NSIZE
    'Algoritmo Marquart
        COEF_MATRIX(i, i) = COEF_MATRIX(i, i) * (1 + TEMP_MULT)
    Next i
    
    TEMP_ARR = MATRIX_TRIANGULAR_LINEAR_SYSTEM_FUNC(COEF_MATRIX)
    If TEMP_ARR(UBound(TEMP_ARR)) = 0 Then: GoTo ERROR_LABEL 'Determ = 0
    DELTA_VECTOR = TEMP_ARR(LBound(TEMP_ARR))
     
    For i = 1 To NSIZE
        DELTA_VECTOR(i, 1) = PARAM_VECTOR(i, 1) + _
                             DELTA_VECTOR(i, 1)
    Next i
        
    FIT_VECTOR = LEVENBERG_MARQUARDT_OBJECTIVE_FUNC(XDATA_MATRIX, _
                      DELTA_VECTOR, FUNC_NAME_STR)
    
    TEMP_SUM = 0
    For i = 1 To NROWS
        TEMP_SUM = TEMP_SUM + _
            (YDATA_VECTOR(i, 1) - FIT_VECTOR(i, 1)) ^ 2
    Next i
    
    BTEMP_VALUE = TEMP_SUM

    If BTEMP_VALUE > tolerance Then ATEMP_DELTA = BTEMP_VALUE / ATEMP_VALUE
        
    BTEMP_DELTA = (BTEMP_VALUE - ATEMP_VALUE)
    If BTEMP_VALUE > tolerance Then: BTEMP_DELTA = BTEMP_DELTA / BTEMP_VALUE
    
    If ATEMP_DELTA > 10 Then
            TEMP_MULT = TEMP_MULT * TEMP_FACT
    Else
        If Abs(BTEMP_DELTA) < LOWER_BOUND Then h = h + 1
        TEMP_MULT = TEMP_MULT / TEMP_FACT
        For i = 1 To NSIZE
            PARAM_VECTOR(i, 1) = DELTA_VECTOR(i, 1) 'update new point
        Next i
        ATEMP_VALUE = BTEMP_VALUE
    End If
        
        'check TEMP_MULT oscillation
    If COUNTER Mod 4 = 0 Then
        If MID_BOUND = TEMP_MULT Then
            TEMP_FACT = 2  'relaxed step
        Else
            TEMP_FACT = 10 'fast step
        End If
        MID_BOUND = TEMP_MULT
    End If
        
    COUNTER = COUNTER + 1
Loop
    
'----------------------------------------------------------------------------------------
Select Case OUTPUT
'----------------------------------------------------------------------------------------
Case 0
    LEVENBERG_MARQUARDT_OPTIMIZATION2_FUNC = PARAM_VECTOR 'Final Parameters
'----------------------------------------------------------------------------------------
Case 1
    
    LEVENBERG_MARQUARDT_OPTIMIZATION2_FUNC = LEVENBERG_MARQUARDT_OBJECTIVE_FUNC(XDATA_MATRIX, _
         PARAM_VECTOR, FUNC_NAME_STR)

'----------------------------------------------------------------------------------------
Case 2
    LEVENBERG_MARQUARDT_OPTIMIZATION2_FUNC = LEVENBERG_MARQUARDT_GRAD_FUNC(XDATA_MATRIX, _
              PARAM_VECTOR, GRAD_STR_NAME)
'----------------------------------------------------------------------------------------
Case 3
    LEVENBERG_MARQUARDT_OPTIMIZATION2_FUNC = LEVENBERG_MARQUARDT_GRAD_FD_VALID_FUNC(XDATA_RNG, _
              PARAM_VECTOR, FUNC_NAME_STR, GRAD_STR_NAME, _
              tolerance, epsilon)
'----------------------------------------------------------------------------------------
Case 4
    LEVENBERG_MARQUARDT_OPTIMIZATION2_FUNC = LEVENBERG_MARQUARDT_SIGMA_FUNC(XDATA_RNG, _
                              YDATA_RNG, PARAM_VECTOR, FUNC_NAME_STR, _
                              GRAD_STR_NAME, 0, epsilon)

'----------------------------------------------------------------------------------------
Case Else
    LEVENBERG_MARQUARDT_OPTIMIZATION2_FUNC = COUNTER
'----------------------------------------------------------------------------------------
End Select
'----------------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
LEVENBERG_MARQUARDT_OPTIMIZATION2_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : LEVENBERG_MARQUARDT_OPTIMIZATION3_FUNC

'DESCRIPTION   : This is a popular alternative to the Gauss-Newton
'method of finding the minimum of a function that is a sum of squares
'of nonlinear functions. This algorithm was found to be an efficient,
'fast and robust method which also has a good global convergence
'property.

'This function has several routines for performing least squares fitting of
'nonlinear functions directly on the worksheet with the Levenberg-Marquardt
'algorithm3. It uses the derivatives information (if available) or approximates
'them internally by the finite difference method. It needs also the function
'definition cell (=f(x, p1, p2,...), the parameter starting values
'(p1, p2,...), and of course the dataset (xi, yi).

'The Levenberg-Marquardt method uses a search direction that is a cross between
'the Gauss-Newton direction and the steepest descent. It is an heuristic method
'that works extremely well in practice. It has become a virtual standard for
'optimization of medium sized nonlinear models.

' REFERENCES:
' Numerical Recipies in Fortran77; W.H. Press, et al.; Cambridge U. Press
' Pages 683 - 685
' Metodos numericos con Matlab; J. M Mathewss et al.; Prentice Hall

'LIBRARY       : OPTIMIZATION
'GROUP         : NLR_LMA
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function LEVENBERG_MARQUARDT_OPTIMIZATION3_FUNC(ByRef XDATA_RNG As Variant, _
ByRef YDATA_RNG As Variant, _
ByRef PARAM_RNG As Variant, _
ByVal FUNC_NAME_STR As String, _
ByVal GRAD_STR_NAME As String, _
Optional ByVal nLOOPS As Variant = 1000, _
Optional ByRef LOWER_BOUND As Variant = 10 ^ 3, _
Optional ByRef UPPER_BOUND As Variant = 10 ^ 9, _
Optional ByVal tolerance As Double = 10 ^ -15)

Dim i As Long

Dim NROWS As Long
Dim NCOLUMNS As Long
Dim NSIZE As Long
Dim COUNTER As Long

Dim TEMP_SUM As Double
Dim TEMP_RESID As Double
Dim TEMP_MULT As Double
Dim TEMP_FACTOR As Double

Dim XDATA_MATRIX As Variant
Dim YDATA_VECTOR As Variant
Dim PARAM_VECTOR As Variant

Dim DELTA_VECTOR As Variant
Dim BETA_VECTOR As Variant
Dim COEF_MATRIX As Variant

Dim FIT_VECTOR As Variant
Dim GRAD_MATRIX As Variant

On Error GoTo ERROR_LABEL

XDATA_MATRIX = XDATA_RNG
NROWS = UBound(XDATA_MATRIX, 1)
NCOLUMNS = UBound(XDATA_MATRIX, 2)

YDATA_VECTOR = YDATA_RNG
If UBound(YDATA_VECTOR, 1) = 1 Then: YDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(YDATA_VECTOR)
If UBound(YDATA_VECTOR, 1) <> NROWS Then: GoTo ERROR_LABEL

PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)
NSIZE = UBound(PARAM_VECTOR, 1)
    

TEMP_FACTOR = 10

TEMP_MULT = LOWER_BOUND

COUNTER = 1

ReDim DELTA_VECTOR(1 To NSIZE, 1 To 1)

Do While COUNTER <= nLOOPS

    FIT_VECTOR = LEVENBERG_MARQUARDT_OBJECTIVE_FUNC(XDATA_MATRIX, _
                      PARAM_VECTOR, FUNC_NAME_STR)
                      
    FIT_VECTOR = MATRIX_ELEMENTS_SUBTRACT_FUNC(YDATA_VECTOR, FIT_VECTOR, 1, 1)

    TEMP_RESID = MATRIX_SUM_PRODUCT_FUNC(FIT_VECTOR, FIT_VECTOR) ^ 0.5

    
    GRAD_MATRIX = LEVENBERG_MARQUARDT_GRAD_FUNC(XDATA_MATRIX, PARAM_VECTOR, _
                      GRAD_STR_NAME)

    BETA_VECTOR = MMULT_FUNC(MATRIX_TRANSPOSE_FUNC(GRAD_MATRIX), FIT_VECTOR, 70)

    COEF_MATRIX = MMULT_FUNC(MATRIX_TRANSPOSE_FUNC(GRAD_MATRIX), GRAD_MATRIX, 70)
    
    For i = 1 To NSIZE
        COEF_MATRIX(i, i) = COEF_MATRIX(i, i) * (1 + 1 / TEMP_MULT)
    Next i
       
    GRAD_MATRIX = MATRIX_LU_LINEAR_SYSTEM_FUNC(COEF_MATRIX, BETA_VECTOR)
    'GRAD_MATRIX = MMULT_FUNC(MATRIX_LU_INVERSE_FUNC(COEF_MATRIX), BETA_VECTOR, 70)
            
    For i = 1 To NSIZE
        DELTA_VECTOR(i, 1) = PARAM_VECTOR(i, 1) + GRAD_MATRIX(i, 1)
    Next i

    FIT_VECTOR = LEVENBERG_MARQUARDT_OBJECTIVE_FUNC(XDATA_MATRIX, DELTA_VECTOR, FUNC_NAME_STR)
    
    FIT_VECTOR = MATRIX_ELEMENTS_SUBTRACT_FUNC(YDATA_VECTOR, FIT_VECTOR, 1, 1)
    

    TEMP_SUM = MATRIX_SUM_PRODUCT_FUNC(FIT_VECTOR, FIT_VECTOR) ^ 0.5

    If (TEMP_RESID <= TEMP_SUM) Then
        TEMP_MULT = TEMP_MULT / TEMP_FACTOR
    Else
        TEMP_MULT = TEMP_MULT * TEMP_FACTOR
        For i = 1 To NSIZE
            PARAM_VECTOR(i, 1) = DELTA_VECTOR(i, 1)
        Next i
    End If

    If Abs(TEMP_SUM - UPPER_BOUND) < tolerance Then: Exit Do
    UPPER_BOUND = TEMP_SUM
    
    COUNTER = COUNTER + 1
Loop

LEVENBERG_MARQUARDT_OPTIMIZATION3_FUNC = PARAM_VECTOR 'Final Parameters

Exit Function
ERROR_LABEL:
LEVENBERG_MARQUARDT_OPTIMIZATION3_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : LEVENBERG_MARQUARDT_OPTIMIZATION4_FUNC

'DESCRIPTION   : The Levenberg-Marquardt algorithm provides a numerical solution
'to the mathematical problem of minimizing a function, generally nonlinear, over
'a space of parameters of the function

'LIBRARY       : OPTIMIZATION
'GROUP         : NLR_LMA
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function LEVENBERG_MARQUARDT_OPTIMIZATION4_FUNC(ByVal FUNC_NAME_STR As String, _
ByRef XDATA_RNG As Variant, _
ByRef YDATA_RNG As Variant, _
ByRef PARAM_RNG As Variant, _
Optional ByRef SCALE_RNG As Variant, _
Optional ByRef OPTION_RNG As Variant, _
Optional ByRef EPS_RNG As Variant, _
Optional ByVal nLOOPS As Long = 1000, _
Optional ByVal tolerance As Double = 0.001, _
Optional ByVal epsilon As Double = 2 ^ 52)
  
Dim i As Long
Dim j As Long
Dim k As Long

Dim NROWS As Long
  
Dim TEMP_SUM As Double
Dim TEMP_GOAL As Double
Dim TEMP_MAX As Double
Dim TEMP_DELTA As Double
Dim TEMP_VALUE As Double

Dim XDATA_MATRIX As Variant
Dim YDATA_VECTOR As Variant

Dim XTEMP_VECTOR As Variant
Dim YTEMP_VECTOR As Variant

Dim PARAM_VECTOR As Variant

Dim TEMP_ARR As Variant

Dim ABS_VECTOR As Variant
Dim DELTA_VECTOR As Variant
Dim DIAGON_VECTOR As Variant
Dim OPTION_VECTOR As Variant
Dim FUNCTION_VECTOR As Variant

Dim ATEMP_VECTOR As Variant
Dim BTEMP_VECTOR As Variant
Dim CTEMP_VECTOR As Variant
Dim DTEMP_VECTOR As Variant
Dim ETEMP_VECTOR As Variant

Dim ATEMP_MATRIX As Variant
Dim BTEMP_MATRIX As Variant

Dim GRAD_MATRIX As Variant

Dim SCALE_VECTOR As Variant
Dim EPSILON_VECTOR As Variant

Dim upper_limit As Variant
Dim lower_limit As Variant
Dim MID_LIMIT As Variant

Dim LAMBDA_VECTOR As Variant

On Error GoTo ERROR_LABEL

MID_LIMIT = 0.0000001
upper_limit = epsilon
lower_limit = 1 / upper_limit

ReDim LAMBDA_VECTOR(1 To 5, 1 To 1)

LAMBDA_VECTOR(1, 1) = 0.1
LAMBDA_VECTOR(2, 1) = 1
LAMBDA_VECTOR(3, 1) = 100#
LAMBDA_VECTOR(4, 1) = 10000#
LAMBDA_VECTOR(5, 1) = 1000000#

XDATA_MATRIX = XDATA_RNG
    If UBound(XDATA_MATRIX, 1) = 1 Then: XDATA_MATRIX = MATRIX_TRANSPOSE_FUNC(XDATA_MATRIX)

YDATA_VECTOR = YDATA_RNG
    If UBound(YDATA_VECTOR, 1) = 1 Then: YDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(YDATA_VECTOR)

PARAM_VECTOR = PARAM_RNG
    If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)

EPSILON_VECTOR = EPS_RNG
If IsArray(EPSILON_VECTOR) = False Or UBound(EPSILON_VECTOR, 1) = 0 Then: _
  EPSILON_VECTOR = _
      MATRIX_ELEMENTS_MULT_SCALAR_FUNC(MATRIX_ELEMENTS_ADD_SCALAR_FUNC( _
      MATRIX_ELEMENTS_MULT_SCALAR_FUNC(PARAM_VECTOR, 0), 1), tolerance)

SCALE_VECTOR = SCALE_RNG
If IsArray(SCALE_VECTOR) = False Or UBound(SCALE_VECTOR, 1) = 0 Then: _
  SCALE_VECTOR = MATRIX_GENERATOR_FUNC(UBound(YDATA_VECTOR, 1), 1, 1)

NROWS = UBound(PARAM_VECTOR, 1)

If UBound(XDATA_MATRIX, 1) <> UBound(YDATA_VECTOR, 1) Then: _
  GoTo ERROR_LABEL 'input(XDATA_MATRIX)/output(YDATA_VECTOR) _
  data must have same number of rows

OPTION_VECTOR = OPTION_RNG
If IsArray(OPTION_VECTOR) = False Then
  OPTION_VECTOR = MATRIX_GENERATOR_FUNC(NROWS, 2, 0)
  
  BTEMP_MATRIX = MATRIX_GENERATOR_FUNC(NROWS, 1, 0)
  For i = 1 To UBound(OPTION_VECTOR, 1)
      OPTION_VECTOR(i, 1) = BTEMP_MATRIX(i, 1)
  Next i

  BTEMP_MATRIX = MATRIX_ELEMENTS_MULT_SCALAR_FUNC(MATRIX_GENERATOR_FUNC(NROWS, 1, 1), _
      upper_limit)
  For i = 1 To UBound(OPTION_VECTOR, 1)
      OPTION_VECTOR(i, 2) = BTEMP_MATRIX(i, 1)
  Next i

Else
  If UBound(OPTION_VECTOR, 1) <> NROWS Then: GoTo ERROR_LABEL
    'OPTION_VECTOR and parameter matrices must have same number of rows
End If

ATEMP_VECTOR = MATRIX_GET_COLUMN_FUNC(OPTION_VECTOR, 1, 1)
BTEMP_VECTOR = MATRIX_GET_COLUMN_FUNC(OPTION_VECTOR, 2, 1)

FUNCTION_VECTOR = Excel.Application.Run(FUNC_NAME_STR, XDATA_MATRIX, PARAM_VECTOR)
YTEMP_VECTOR = FUNCTION_VECTOR

XTEMP_VECTOR = PARAM_VECTOR

ATEMP_MATRIX = MATRIX_ELEMENTS_MULT_FUNC(SCALE_VECTOR, MATRIX_ELEMENTS_SUBTRACT_FUNC(YDATA_VECTOR, _
  FUNCTION_VECTOR, 1, 1), 1, 1)

TEMP_VALUE = VECTOR_ELEMENTS_DOT_PRODUCT_FUNC(ATEMP_MATRIX, ATEMP_MATRIX)

CTEMP_VECTOR = MATRIX_GENERATOR_FUNC(NROWS, 1, 0)
ETEMP_VECTOR = MATRIX_ELEMENTS_MULT_SCALAR_FUNC(MATRIX_GENERATOR_FUNC(NROWS, 1, 1), upper_limit)
TEMP_DELTA = 1

'----------------------------------------------------------------------------------------
For i = 1 To nLOOPS
'----------------------------------------------------------------------------------------
  
  GRAD_MATRIX = JACOBI_MATRIX_FUNC(FUNC_NAME_STR, XDATA_MATRIX, _
                XTEMP_VECTOR, EPSILON_VECTOR)
  DTEMP_VECTOR = XTEMP_VECTOR
  
  ATEMP_MATRIX = MATRIX_ELEMENTS_MULT_FUNC(SCALE_VECTOR, _
      MATRIX_ELEMENTS_SUBTRACT_FUNC(YDATA_VECTOR, YTEMP_VECTOR, 1, 1), 1, 1)
  
  TEMP_GOAL = (1 - tolerance) * TEMP_VALUE
  
  For j = 1 To NROWS
    If (EPSILON_VECTOR(j, 1) = 0) Then
      CTEMP_VECTOR(j, 1) = 0
    Else
      
      BTEMP_MATRIX = MATRIX_ELEMENTS_MULT_FUNC(SCALE_VECTOR, _
          MATRIX_GET_COLUMN_FUNC(GRAD_MATRIX, j, 1), 1, 1)
      For k = 1 To UBound(GRAD_MATRIX, 1)
          GRAD_MATRIX(k, j) = BTEMP_MATRIX(k, 1)
      Next k
      
      CTEMP_VECTOR(j, 1) = MATRIX_SUM_PRODUCT_FUNC(MATRIX_GET_COLUMN_FUNC(GRAD_MATRIX, j, 1), _
      MATRIX_GET_COLUMN_FUNC(GRAD_MATRIX, j, 1))
      
      If (CTEMP_VECTOR(j, 1) > 0) Then: _
          CTEMP_VECTOR(j, 1) = 1 / (CTEMP_VECTOR(j, 1) ^ 0.5)
      
     
      BTEMP_MATRIX = MATRIX_ELEMENTS_MULT_SCALAR_FUNC(MATRIX_GET_COLUMN_FUNC(GRAD_MATRIX, j, 1), _
          CTEMP_VECTOR(j, 1))
      For k = 1 To UBound(GRAD_MATRIX, 1)
          GRAD_MATRIX(k, j) = BTEMP_MATRIX(k, 1)
      Next k
    End If
  Next j

'-------------------------------------------------------------------------------
  TEMP_ARR = MATRIX_SVD_FACT_FUNC(GRAD_MATRIX, 0)
  GRAD_MATRIX = MATRIX_ELEMENTS_MULT_SCALAR_FUNC(TEMP_ARR(1), -1)
  BTEMP_MATRIX = MATRIX_ELEMENTS_MULT_SCALAR_FUNC(TEMP_ARR(2), -1)
  DIAGON_VECTOR = MATRIX_DIAGONAL_VECTOR_FUNC(TEMP_ARR(3), 1)
'-------------------------------------------------------------------------------
  
  For j = 1 To UBound(LAMBDA_VECTOR, 1)
    
    TEMP_MAX = MAXIMUM_FUNC(TEMP_DELTA * LAMBDA_VECTOR(j, 1), MID_LIMIT)
    
    DELTA_VECTOR = MATRIX_ELEMENTS_MULT_FUNC(MMULT_FUNC(BTEMP_MATRIX, _
      MATRIX_ELEMENTS_DIVIDE_FUNC(MMULT_FUNC(MATRIX_TRANSPOSE_FUNC(GRAD_MATRIX), ATEMP_MATRIX), _
      MATRIX_ELEMENTS_POWER_FUNC(MATRIX_ELEMENTS_ADD_SCALAR_FUNC(MATRIX_ELEMENTS_MULT_FUNC(DIAGON_VECTOR, _
      DIAGON_VECTOR, 1, 1), TEMP_MAX), 0.5, 0), 1, 1)), CTEMP_VECTOR, 1, 1) _
    'check the change constraints and apply As necessary
    
    For k = 1 To NROWS
      If BTEMP_VECTOR(k, 1) = upper_limit Then: Exit For
      
      DELTA_VECTOR(k, 1) = MAXIMUM_FUNC(DELTA_VECTOR(k, 1), _
          -Abs(BTEMP_VECTOR(k, 1) * DTEMP_VECTOR(k, 1)))
      
      DELTA_VECTOR(k, 1) = MINIMUM_FUNC(DELTA_VECTOR(k, 1), _
          Abs(BTEMP_VECTOR(k, 1) * DTEMP_VECTOR(k, 1)))
    Next k
    
    ABS_VECTOR = _
      MATRIX_ABSOLUTE_FUNC(MATRIX_ELEMENTS_MULT_FUNC(ATEMP_VECTOR, XTEMP_VECTOR, 1, 1))
    
    If MATRIX_CHECK_VALUE_FUNC(MATRIX_ELEMENTS_GREATER_FUNC(MATRIX_ABSOLUTE_FUNC(DELTA_VECTOR), _
      MATRIX_ELEMENTS_MULT_SCALAR_FUNC(ABS_VECTOR, LAMBDA_VECTOR(1, 1))), 0, 0) Then
      
      PARAM_VECTOR = MATRIX_ELEMENTS_ADD_FUNC(DELTA_VECTOR, DTEMP_VECTOR, 1, 1)
      FUNCTION_VECTOR = Excel.Application.Run(FUNC_NAME_STR, XDATA_MATRIX, PARAM_VECTOR)
      
      ATEMP_MATRIX = MATRIX_ELEMENTS_MULT_FUNC(SCALE_VECTOR, _
          MATRIX_ELEMENTS_SUBTRACT_FUNC(YDATA_VECTOR, FUNCTION_VECTOR, 1, 1), 1, 1)
      
      TEMP_SUM = MATRIX_SUM_PRODUCT_FUNC(ATEMP_MATRIX, ATEMP_MATRIX)
      
      If (TEMP_SUM < TEMP_VALUE) Then
        XTEMP_VECTOR = PARAM_VECTOR
        YTEMP_VECTOR = FUNCTION_VECTOR
        TEMP_VALUE = TEMP_SUM
      End If
      
      If TEMP_SUM <= TEMP_GOAL Then: Exit For
    End If 'end if Abs(DELTA_VECTOR)
  Next j 'end for j
  
  TEMP_DELTA = TEMP_MAX
  
  If TEMP_SUM < lower_limit Then: Exit For
      ABS_VECTOR = _
          MATRIX_ABSOLUTE_FUNC(MATRIX_ELEMENTS_MULT_FUNC(ATEMP_VECTOR, XTEMP_VECTOR, 1, 1))
  If (MATRIX_CHECK_VALUE_FUNC(MATRIX_ELEMENTS_LESS_FUNC(MATRIX_ABSOLUTE_FUNC(DELTA_VECTOR), _
          ABS_VECTOR), 0, 1) And _
      MATRIX_CHECK_VALUE_FUNC(MATRIX_ELEMENTS_LESS_FUNC(MATRIX_ABSOLUTE_FUNC(ETEMP_VECTOR), _
          ABS_VECTOR), 0, 1)) Then 'Parameter changes converged to
  'specified precision
    Exit For
  Else
      ETEMP_VECTOR = DELTA_VECTOR
  End If
  
  If (TEMP_SUM > TEMP_GOAL) Then: Exit For
'----------------------------------------------------------------------------------------
Next i
'----------------------------------------------------------------------------------------
  
LEVENBERG_MARQUARDT_OPTIMIZATION4_FUNC = Array(XTEMP_VECTOR, YTEMP_VECTOR)
  
Exit Function
ERROR_LABEL:
LEVENBERG_MARQUARDT_OPTIMIZATION4_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : LEVENBERG_MARQUARDT_OBJECTIVE_FUNC

'DESCRIPTION   : This Objective function was developed for performing the
'optimization task directly on a worksheet.

'This means that you can define any relationship that you want to
'optimize, simply by using the any standard optimization function
'and the range that relate them.

'LIBRARY       : OPTIMIZATION
'GROUP         : NLR_LMA
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************


Function LEVENBERG_MARQUARDT_OBJECTIVE_FUNC(ByRef XDATA_RNG As Variant, _
ByRef PARAM_RNG As Variant, _
ByVal FUNC_NAME_STR As String)

Dim PARAM_VECTOR As Variant
Dim XDATA_MATRIX As Variant

Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

'--------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------
XDATA_MATRIX = XDATA_RNG

PARAM_VECTOR = PARAM_RNG
    If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)
'--------------------------------------------------------------------------------------
TEMP_MATRIX = Excel.Application.Run(FUNC_NAME_STR, XDATA_MATRIX, PARAM_VECTOR)
LEVENBERG_MARQUARDT_OBJECTIVE_FUNC = TEMP_MATRIX
'--------------------------------------------------------------------------------------
Exit Function
ERROR_LABEL:
LEVENBERG_MARQUARDT_OBJECTIVE_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : LEVENBERG_MARQUARDT_GRAD_FUNC

'DESCRIPTION   : This Gradient function was developed for performing the
'optimization task directly on a worksheet.

'This means that you can define any relationship that you want to
'optimize, simply by using the any standard optimization function
'and the range that relate them.

'LIBRARY       : OPTIMIZATION
'GROUP         : NLR_LMA
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************


Function LEVENBERG_MARQUARDT_GRAD_FUNC(ByRef XDATA_RNG As Variant, _
ByRef PARAM_RNG As Variant, _
ByVal GRAD_STR_NAME As String)

Dim XDATA_MATRIX As Variant
Dim PARAM_VECTOR As Variant
Dim GRAD_MATRIX As Variant

On Error GoTo ERROR_LABEL
XDATA_MATRIX = XDATA_RNG

PARAM_VECTOR = PARAM_RNG
    If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)

GRAD_MATRIX = Excel.Application.Run(GRAD_STR_NAME, XDATA_MATRIX, PARAM_VECTOR)
    
LEVENBERG_MARQUARDT_GRAD_FUNC = GRAD_MATRIX

Exit Function
ERROR_LABEL:
LEVENBERG_MARQUARDT_GRAD_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : LEVENBERG_MARQUARDT_GRAD_FD_VALID_FUNC
'DESCRIPTION   : Check the congruence of exact derivative with the
'approximated derivatives in order to avoid derivatives
'formulas errors.
'LIBRARY       : OPTIMIZATION
'GROUP         : NLR_LMA
'ID            : 007
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************


Function LEVENBERG_MARQUARDT_GRAD_FD_VALID_FUNC(ByRef XDATA_RNG As Variant, _
ByRef PARAM_RNG As Variant, _
ByVal FUNC_NAME_STR As String, _
ByVal GRAD_STR_NAME As String, _
Optional ByVal tolerance As Double = 10 ^ -6, _
Optional ByVal epsilon As Double = 10 ^ -4)

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim ATEMP_VALUE As Double
Dim BTEMP_VALUE As Double

Dim XDATA_MATRIX As Variant
Dim TEMP_MATRIX As Variant
Dim GRAD_MATRIX As Variant
Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL

LEVENBERG_MARQUARDT_GRAD_FD_VALID_FUNC = False

'--------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------
XDATA_MATRIX = XDATA_RNG
PARAM_VECTOR = PARAM_RNG
    If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)
'--------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------

GRAD_MATRIX = LEVENBERG_MARQUARDT_GRAD_FUNC(XDATA_MATRIX, PARAM_VECTOR, _
                      GRAD_STR_NAME)


TEMP_MATRIX = LEVENBERG_MARQUARDT_GRAD_FD_APPROX_FUNC(XDATA_MATRIX, PARAM_VECTOR, _
                      FUNC_NAME_STR, epsilon)
'--------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------
    
NROWS = UBound(TEMP_MATRIX, 1)
NCOLUMNS = UBound(TEMP_MATRIX, 2)
    
ATEMP_VALUE = 0
BTEMP_VALUE = 0
For j = 1 To NCOLUMNS
    For i = 1 To NROWS
        ATEMP_VALUE = (Abs(TEMP_MATRIX(i, j)) + Abs(GRAD_MATRIX(i, j)))
        BTEMP_VALUE = Abs(TEMP_MATRIX(i, j) - GRAD_MATRIX(i, j))
        If ATEMP_VALUE > tolerance Then: BTEMP_VALUE = BTEMP_VALUE / ATEMP_VALUE
        If BTEMP_VALUE > epsilon Then
            LEVENBERG_MARQUARDT_GRAD_FD_VALID_FUNC = False 'Derivatives: dubious accuracy.
            Exit Function
        End If
    Next i
Next j

LEVENBERG_MARQUARDT_GRAD_FD_VALID_FUNC = True

Exit Function
ERROR_LABEL:
LEVENBERG_MARQUARDT_GRAD_FD_VALID_FUNC = False
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : LEVENBERG_MARQUARDT_GRAD_FD_APPROX_FUNC
'DESCRIPTION   : Approximate gradients with finite differences
'LIBRARY       : OPTIMIZATION
'GROUP         : NLR_LMA
'ID            : 008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function LEVENBERG_MARQUARDT_GRAD_FD_APPROX_FUNC(ByRef XDATA_RNG As Variant, _
ByRef PARAM_RNG As Variant, _
ByVal FUNC_NAME_STR As String, _
Optional ByVal epsilon As Double = 10 ^ -4)

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim NSIZE As Long

Dim TEMP_VALUE As Double

Dim ATEMP_VECTOR As Variant
Dim BTEMP_VECTOR As Variant
Dim CTEMP_VECTOR As Variant

Dim PARAM_VECTOR As Variant

Dim XDATA_MATRIX As Variant
Dim GRAD_MATRIX As Variant

On Error GoTo ERROR_LABEL

'--------------------------------------------------------------------------------------
XDATA_MATRIX = XDATA_RNG
NROWS = UBound(XDATA_MATRIX, 1)
'--------------------------------------------------------------------------------------

PARAM_VECTOR = PARAM_RNG
    If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)
NSIZE = UBound(PARAM_VECTOR, 1)
'--------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------
            
ReDim ATEMP_VECTOR(1 To NSIZE, 1 To 1)

For j = 1 To NSIZE 'step forward
    For i = 1 To NSIZE
        If i = j Then
            TEMP_VALUE = PARAM_VECTOR(i, 1) + epsilon / 2
        Else
            TEMP_VALUE = PARAM_VECTOR(i, 1)
        End If
        ATEMP_VECTOR(i, 1) = TEMP_VALUE
    Next i
                
    BTEMP_VECTOR = LEVENBERG_MARQUARDT_OBJECTIVE_FUNC(XDATA_MATRIX, _
                   ATEMP_VECTOR, FUNC_NAME_STR)

    For i = 1 To NSIZE 'step back
        If i = j Then
            TEMP_VALUE = PARAM_VECTOR(i, 1) - epsilon / 2
        Else
            TEMP_VALUE = PARAM_VECTOR(i, 1)
        End If
        ATEMP_VECTOR(i, 1) = TEMP_VALUE
    Next i
                
    CTEMP_VECTOR = LEVENBERG_MARQUARDT_OBJECTIVE_FUNC(XDATA_MATRIX, _
                   ATEMP_VECTOR, FUNC_NAME_STR)
                
    NROWS = UBound(BTEMP_VECTOR, 1) 'finite difference
    If j = 1 Then: ReDim GRAD_MATRIX(1 To NROWS, 1 To NSIZE)
    For i = 1 To NROWS
        GRAD_MATRIX(i, j) = (BTEMP_VECTOR(i, 1) - _
                           CTEMP_VECTOR(i, 1)) / epsilon
    Next i
Next j


LEVENBERG_MARQUARDT_GRAD_FD_APPROX_FUNC = GRAD_MATRIX

Exit Function
ERROR_LABEL:
LEVENBERG_MARQUARDT_GRAD_FD_APPROX_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : LEVENBERG_MARQUARDT_SIGMA_FUNC
'DESCRIPTION   : Compute the standard deviation of the estimates of LM fitting
'LIBRARY       : OPTIMIZATION
'GROUP         : NLR_LMA
'ID            : 009
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function LEVENBERG_MARQUARDT_SIGMA_FUNC(ByRef XDATA_RNG As Variant, _
ByRef YDATA_RNG As Variant, _
ByRef PARAM_RNG As Variant, _
ByVal FUNC_NAME_STR As String, _
Optional ByVal GRAD_STR_NAME As String = "", _
Optional ByVal OUTPUT As Integer = 1, _
Optional ByVal epsilon As Double = 10 ^ -4)

Dim i As Long

Dim NSIZE As Long
Dim NROWS As Long

Dim TEMP_SUM As Double
Dim FIT_VECTOR As Variant
Dim YDATA_VECTOR As Variant
Dim PARAM_VECTOR As Variant

Dim GRAD_MATRIX As Variant
Dim XDATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

XDATA_MATRIX = XDATA_RNG
NROWS = UBound(XDATA_MATRIX, 1)

YDATA_VECTOR = YDATA_RNG
If UBound(YDATA_VECTOR, 1) = 1 Then: YDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(YDATA_VECTOR)
If UBound(YDATA_VECTOR, 1) <> UBound(XDATA_MATRIX, 1) Then: GoTo ERROR_LABEL

PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)
NSIZE = UBound(PARAM_VECTOR, 1)

'----------------------------------------------------------------------------------
FIT_VECTOR = LEVENBERG_MARQUARDT_OBJECTIVE_FUNC(XDATA_MATRIX, PARAM_VECTOR, FUNC_NAME_STR)
'----------------------------------------------------------------------------------

'----------------------------------------------------------------------------------
If GRAD_STR_NAME = "" Then
'----------------------------------------------------------------------------------
    GRAD_MATRIX = LEVENBERG_MARQUARDT_GRAD_FD_APPROX_FUNC(XDATA_MATRIX, PARAM_VECTOR, _
                      FUNC_NAME_STR, epsilon)
'----------------------------------------------------------------------------------
Else
'----------------------------------------------------------------------------------
    GRAD_MATRIX = LEVENBERG_MARQUARDT_GRAD_FUNC(XDATA_MATRIX, PARAM_VECTOR, _
                      GRAD_STR_NAME)
'----------------------------------------------------------------------------------
End If
'----------------------------------------------------------------------------------
        
GRAD_MATRIX = MATRIX_LU_INVERSE_FUNC(MMULT_FUNC(MATRIX_TRANSPOSE_FUNC(GRAD_MATRIX), _
        GRAD_MATRIX, 70)) ' Compute the inverse of the Jacobian


TEMP_SUM = 0
For i = 1 To NROWS
    TEMP_SUM = TEMP_SUM + (YDATA_VECTOR(i, 1) - FIT_VECTOR(i, 1)) ^ 2
Next i
TEMP_SUM = Sqr(TEMP_SUM / (NROWS - NSIZE)) 'residual standard deviation _
                   of the LM fitting

If NSIZE > 1 Then
    ReDim PARAM_VECTOR(1 To NSIZE, 1 To 1)
    For i = 1 To NSIZE
        PARAM_VECTOR(i, 1) = Sqr(GRAD_MATRIX(i, i)) * _
                             TEMP_SUM
    Next i
Else
    ReDim PARAM_VECTOR(1 To 1, 1 To 1)
    PARAM_VECTOR(1, 1) = Sqr(GRAD_MATRIX(1, 1)) * TEMP_SUM
End If
        
'---------------------------------------------------------------------------------------
Select Case OUTPUT
'---------------------------------------------------------------------------------------
    Case 0 'Compute the standard deviation of the estimates of LM fitting
        LEVENBERG_MARQUARDT_SIGMA_FUNC = PARAM_VECTOR
'---------------------------------------------------------------------------------------
    Case Else 'Compute the residual standard deviation of the LM fitting
        LEVENBERG_MARQUARDT_SIGMA_FUNC = TEMP_SUM
'---------------------------------------------------------------------------------------
End Select
'---------------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
    LEVENBERG_MARQUARDT_SIGMA_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : LEVENBERG_MARQUARDT_LRE_FUNC
'DESCRIPTION   : LOG RELATIVE ERROR
'LIBRARY       : OPTIMIZATION
'GROUP         : NLR_LMA
'ID            : 010
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function LEVENBERG_MARQUARDT_LRE_FUNC(ByVal X_VAL As Double, _
ByVal X_EST As Double)
'X_VAL --> Certified Value
'X_EST --> Approximated Value
On Error GoTo ERROR_LABEL
LEVENBERG_MARQUARDT_LRE_FUNC = Abs((X_EST - X_VAL) / X_VAL)
Exit Function
ERROR_LABEL:
LEVENBERG_MARQUARDT_LRE_FUNC = Err.number
End Function
