Attribute VB_Name = "OPTIM_NLR_FRAMES_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

Private PUB_LEVENBERG_MARQUARDT_VERS As Integer
Private PUB_RATION_NUMER As Long
Private PUB_RATION_DENOM As Long

'************************************************************************************
'************************************************************************************
'FUNCTION      : LEVENBERG_MARQUARDT_GENERIC_FUNC
'DESCRIPTION   : Nonlinear Regression with a predefined Objective model
'This set of functions comes in handy when we have to perform a nonlinear
'regression using a predefined model. They are much faster that the general
'nonlinear regression and, in addition, you do not have to build the formula
'model and its derivatives.

'LIBRARY       : OPTIMIZATION
'GROUP         : NLR_FRAMES
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function LEVENBERG_MARQUARDT_GENERIC_FUNC(ByRef XDATA_RNG As Variant, _
ByRef YDATA_RNG As Variant, _
ByRef PARAM_RNG As Variant, _
Optional ByVal VERSION As Integer = 4, _
Optional ByVal nLOOPS As Long = 1000, _
Optional ByVal tolerance As Double = 10 ^ -10, _
Optional ByVal epsilon As Double = 10 ^ -6, _
Optional ByVal OUTPUT As Integer = 0)

Dim NROWS As Long

Dim PARAM_VECTOR As Variant
Dim XDATA_MATRIX As Variant
Dim YDATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

PUB_LEVENBERG_MARQUARDT_VERS = VERSION
XDATA_MATRIX = XDATA_RNG
If UBound(XDATA_MATRIX, 1) = 1 Then: XDATA_MATRIX = MATRIX_TRANSPOSE_FUNC(XDATA_MATRIX)

YDATA_VECTOR = YDATA_RNG
If UBound(YDATA_VECTOR, 1) = 1 Then: YDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(YDATA_VECTOR)

If UBound(XDATA_MATRIX, 1) <> UBound(YDATA_VECTOR, 1) Then: GoTo ERROR_LABEL
NROWS = UBound(XDATA_MATRIX, 1)

PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)

'--------------------------------------------------------------------------------
Select Case OUTPUT
'--------------------------------------------------------------------------------
Case 0
'--------------------------------------------------------------------------------
    LEVENBERG_MARQUARDT_GENERIC_FUNC = _
        LEVENBERG_MARQUARDT_OPTIMIZATION2_FUNC(XDATA_MATRIX, YDATA_VECTOR, _
                PARAM_VECTOR, "CALL_GENERIC_OBJ_FUNC", _
                "", 0, nLOOPS, tolerance, epsilon)
'--------------------------------------------------------------------------------
Case 1
'--------------------------------------------------------------------------------
    LEVENBERG_MARQUARDT_GENERIC_FUNC = _
        LEVENBERG_MARQUARDT_OPTIMIZATION2_FUNC(XDATA_MATRIX, YDATA_VECTOR, _
                PARAM_VECTOR, "CALL_GENERIC_OBJ_FUNC", _
                "", 1, nLOOPS, tolerance, epsilon)

'--------------------------------------------------------------------------------
Case 2
'--------------------------------------------------------------------------------
    LEVENBERG_MARQUARDT_GENERIC_FUNC = _
        LEVENBERG_MARQUARDT_OPTIMIZATION2_FUNC(XDATA_MATRIX, YDATA_VECTOR, _
                PARAM_VECTOR, "CALL_GENERIC_OBJ_FUNC", _
                "", 2, nLOOPS, tolerance, epsilon)

'--------------------------------------------------------------------------------
Case 3
'--------------------------------------------------------------------------------
    LEVENBERG_MARQUARDT_GENERIC_FUNC = _
        LEVENBERG_MARQUARDT_OPTIMIZATION2_FUNC(XDATA_MATRIX, YDATA_VECTOR, _
                PARAM_VECTOR, "CALL_GENERIC_OBJ_FUNC", _
                "", 3, nLOOPS, tolerance, epsilon)

'--------------------------------------------------------------------------------
Case 4
'--------------------------------------------------------------------------------
    LEVENBERG_MARQUARDT_GENERIC_FUNC = _
        LEVENBERG_MARQUARDT_OPTIMIZATION2_FUNC(XDATA_MATRIX, YDATA_VECTOR, _
                PARAM_VECTOR, "CALL_GENERIC_OBJ_FUNC", _
                "", 4, nLOOPS, tolerance, epsilon)

'--------------------------------------------------------------------------------
Case Else
'--------------------------------------------------------------------------------
    LEVENBERG_MARQUARDT_GENERIC_FUNC = _
        LEVENBERG_MARQUARDT_OPTIMIZATION2_FUNC(XDATA_MATRIX, YDATA_VECTOR, _
                PARAM_VECTOR, "CALL_GENERIC_OBJ_FUNC", _
                "", 5, nLOOPS, tolerance, epsilon)

'--------------------------------------------------------------------------------
End Select
'--------------------------------------------------------------------------------


Exit Function
ERROR_LABEL:
LEVENBERG_MARQUARDT_GENERIC_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_GENERIC_OBJ_FUNC
'DESCRIPTION   :
'LIBRARY       : OPTIMIZATION
'GROUP         : NLR_FRAMES
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function CALL_GENERIC_OBJ_FUNC(ByRef XDATA_RNG As Variant, _
ByRef PARAM_RNG As Variant)

Dim i As Long
Dim NROWS As Long
Dim TEMP_VECTOR As Variant
Dim PARAM_VECTOR As Variant
Dim XDATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

XDATA_MATRIX = XDATA_RNG
NROWS = UBound(XDATA_MATRIX, 1)
PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)

'------------------------------------------------------------------------------------
Select Case PUB_LEVENBERG_MARQUARDT_VERS
'----------------------- -------------------------------------------------------------
Case 0 'multivariable regression A
'------------------------------------------------------------------------------------
    ReDim TEMP_VECTOR(1 To NROWS, 1 To 1)
    For i = 1 To NROWS
        TEMP_VECTOR(i, 1) = 1 / (PARAM_VECTOR(1, 1) * XDATA_MATRIX(i, 1) ^ 2 + _
                PARAM_VECTOR(2, 1) * XDATA_MATRIX(i, 1) * XDATA_MATRIX(i, 2) + _
                PARAM_VECTOR(3, 1) * XDATA_MATRIX(i, 2) ^ 1)
    Next i
'------------------------------------------------------------------------------------
Case 1 'exponential class A
'------------------------------------------------------------------------------------
    ReDim TEMP_VECTOR(1 To NROWS, 1 To 1)
    For i = 1 To NROWS
        TEMP_VECTOR(i, 1) = PARAM_VECTOR(1, 1) * _
                            (1 - Exp(-1 * PARAM_VECTOR(2, 1) * _
                             XDATA_MATRIX(i, 1)))
    Next i
'------------------------------------------------------------------------------------
Case 2 'ArcTan
'------------------------------------------------------------------------------------
    ReDim TEMP_VECTOR(1 To NROWS, 1 To 1)
    For i = 1 To NROWS
        TEMP_VECTOR(i, 1) = Atn(PARAM_VECTOR(1, 1) * XDATA_MATRIX(i, 1) ^ 2) - _
                             Atn(PARAM_VECTOR(2, 1) * XDATA_MATRIX(i, 1) ^ 2)
    Next i
'    ReDim TEMP_VECTOR(1 To NROWS, 1 To 2)
 '   For i = 1 To NROWS
  '      TEMP_VECTOR(i, 1) = XDATA_MATRIX(i, 1) ^ 2 / _
        (PARAM_VECTOR(1, 1) ^ 2 * XDATA_MATRIX(i, 1) ^ 4 + 1)
        
   '     TEMP_VECTOR(i, 2) = -1 * XDATA_MATRIX(i, 1) ^ 2 / _
        (PARAM_VECTOR(2, 1) ^ 2 * XDATA_MATRIX(i, 1) ^ 4 + 1)
    
    'Next i
'------------------------------------------------------------------------------------
Case 3 'rational class
'------------------------------------------------------------------------------------

    ReDim TEMP_VECTOR(1 To NROWS, 1 To 1)
    For i = 1 To NROWS
        TEMP_VECTOR(i, 1) = PARAM_VECTOR(1, 1) * _
        (XDATA_MATRIX(i, 1) ^ 2 + XDATA_MATRIX(i, 1) * _
        PARAM_VECTOR(2, 1)) / (XDATA_MATRIX(i, 1) ^ 2 + XDATA_MATRIX(i, 1) _
        * PARAM_VECTOR(3, 1) + PARAM_VECTOR(4, 1))
    
    Next i
'------------------------------------------------------------------------------------
Case 4 'multi variable regression B
'------------------------------------------------------------------------------------
    ReDim TEMP_VECTOR(1 To NROWS, 1 To 1)
    For i = 1 To NROWS
        TEMP_VECTOR(i, 1) = PARAM_VECTOR(1, 1) * XDATA_MATRIX(i, 1) ^ 2 + _
                             PARAM_VECTOR(2, 1) * XDATA_MATRIX(i, 2) ^ 2 + _
                             PARAM_VECTOR(3, 1) * XDATA_MATRIX(i, 3) ^ 2 - 1
    
    Next i
    'exp(-b1*x)/(b2+b3*x)
'------------------------------------------------------------------------------------
Case 5 'Exponential Class B
'------------------------------------------------------------------------------------
    ReDim TEMP_VECTOR(1 To NROWS, 1 To 1)
    For i = 1 To NROWS
        TEMP_VECTOR(i, 1) = PARAM_VECTOR(1, 1) * Exp(PARAM_VECTOR(2, 1) * _
                             XDATA_MATRIX(i, 1))
    Next i
'------------------------------------------------------------------------------------
Case 6 'Exponential Class C
'------------------------------------------------------------------------------------
    ReDim TEMP_VECTOR(1 To NROWS, 1 To 1)
    For i = 1 To NROWS
        TEMP_VECTOR(i, 1) = Exp(-1 * PARAM_VECTOR(1, 1) * XDATA_MATRIX(i, 1)) / _
                             (PARAM_VECTOR(2, 1) + PARAM_VECTOR(3, 1) * _
                              XDATA_MATRIX(i, 1))
    Next i
'------------------------------------------------------------------------------------
Case Else 'RR,K,A,B --> H1N1 Cases
'------------------------------------------------------------------------------------
    ReDim TEMP_VECTOR(1 To NROWS, 1 To 1)
    For i = 1 To NROWS
        TEMP_VECTOR(i, 1) = PARAM_VECTOR(2, 1) * (1 - PARAM_VECTOR(3, 1) * Exp(-PARAM_VECTOR(1, 1) * _
                            XDATA_MATRIX(i, 1))) ^ PARAM_VECTOR(4, 1)
    Next i
'------------------------------------------------------------------------------------
End Select
'------------------------------------------------------------------------------------

CALL_GENERIC_OBJ_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
CALL_GENERIC_OBJ_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : LEVENBERG_MARQUARDT_RATIONAL_FUNC

'DESCRIPTION   : Rational formulas could be used to approximate a wide
'variety of functions. But the mainly profit happens when we want to
'interpolate a function near a "pole". Large, sharply oscillations of
'the system responce could be followed better using a rational
'model instead other models like polynomials or exponentials.

'A rational model is a fraction of two polynomials. The max degree of
'one polynomial gives the degree of the rational model. Usually, in
'modelling of real stable systems, the denominator degree is always
'greater then the numerator.

'When using the rational model

'The rational model is more complicate then polynomial model. For example
'a 3 degree rational model has 6 parameters, while the polynomial model of
'the same degree has 4 parameters.

'Far from the "pole", the rational model takes no advantage over the
'polynomial model. Therefore they should be used only when it is truly
'necessary. The scatter plot of the dataset can helps us in choosing the
'adapt model. A typical plot that increase or decrease sharply often
'detects the presence of a "pole" and a 1st degree rational model could be
'sufficient; on the other hand, if the plot shows a narrow
'"peak" probably the better choice would be a 2nd degree rational model.

'Usually also the inquiry of the system characteristic from witch the samples
'are given could be help us to choose a suitable rational model degree.

'LIBRARY       : OPTIMIZATION
'GROUP         : NLR_FRAMES
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function LEVENBERG_MARQUARDT_RATIONAL_FUNC(ByRef XDATA_RNG As Variant, _
ByRef YDATA_RNG As Variant, _
ByVal NUMER_DEG As Long, _
ByVal DENOM_DEG As Long, _
Optional ByVal nLOOPS As Variant = 1000, _
Optional ByVal tolerance As Double = 10 ^ -10, _
Optional ByVal epsilon As Double = 10 ^ -6, _
Optional ByVal OUTPUT As Integer = 0)

Dim i As Long
Dim j As Long
Dim k As Long

Dim NROWS As Long
Dim NSIZE As Long

Dim XDATA_VECTOR As Variant
Dim YDATA_VECTOR As Variant

Dim ATEMP_MATRIX As Variant
Dim BTEMP_MATRIX As Variant

Dim NUMER_VECTOR As Variant
Dim DENOM_VECTOR As Variant

Dim XTEMP_VECTOR As Variant
Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL

PUB_RATION_NUMER = 0
PUB_RATION_DENOM = 0

'--------------------------------------------------------------------
'return the coefficient of the rational regression
' (a0+a1*x+a2*x^2+...+an*x^n)/(b0+b1*x+b2*x^2+...+x^NSIZE)
'NUMER_VECTOR = coefficients
'DENOM_VECTOR = numerator coefficients
'NUMER_DEG = max degree of the numerator
'DENOM_DEG = degree of the denominator
'-------------------------------------------------------------------

XDATA_VECTOR = XDATA_RNG
If UBound(XDATA_VECTOR, 1) = 1 Then: XDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(XDATA_VECTOR)
YDATA_VECTOR = YDATA_RNG
If UBound(YDATA_VECTOR, 1) = 1 Then: YDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(YDATA_VECTOR)

If UBound(XDATA_VECTOR, 1) <> UBound(YDATA_VECTOR, 1) Then: GoTo ERROR_LABEL

NROWS = UBound(XDATA_VECTOR, 1)
NSIZE = NUMER_DEG + DENOM_DEG + 1

ReDim PARAM_VECTOR(1 To NROWS, 1 To 1)
ReDim ATEMP_MATRIX(1 To NROWS, 1 To NSIZE)

For j = 1 To NUMER_DEG + 1
    For i = 1 To NROWS
        ATEMP_MATRIX(i, j) = XDATA_VECTOR(i, 1) ^ (j - 1)
    Next i
Next j


For j = 1 To DENOM_DEG
    For i = 1 To NROWS
        ATEMP_MATRIX(i, j + NUMER_DEG + 1) = _
            -XDATA_VECTOR(i, 1) ^ (j - 1) * YDATA_VECTOR(i, 1)
    Next i
Next j

For i = 1 To NROWS
    PARAM_VECTOR(i, 1) = YDATA_VECTOR(i, 1) * XDATA_VECTOR(i, 1) ^ DENOM_DEG
Next i


'ReDim BTEMP_MATRIX(1 To NSIZE, 1 To NSIZE)
ReDim BTEMP_MATRIX(1 To NSIZE, 1 To NSIZE + 1)

For i = 1 To NSIZE
    For j = 1 To NSIZE
        For k = 1 To NROWS
            BTEMP_MATRIX(i, j) = BTEMP_MATRIX(i, j) + ATEMP_MATRIX(k, i) * _
                                 ATEMP_MATRIX(k, j)
        Next k
    Next j
Next i

'j = 1
j = NSIZE + 1
ReDim XTEMP_VECTOR(1 To NSIZE, 1 To j)
For i = 1 To NSIZE
    For k = 1 To NROWS
        
'        XTEMP_VECTOR(i, j) = XTEMP_VECTOR(i, j) + ATEMP_MATRIX(k, i) * _
                             PARAM_VECTOR(k, 1)
        
        BTEMP_MATRIX(i, j) = BTEMP_MATRIX(i, j) + ATEMP_MATRIX(k, i) * _
                             PARAM_VECTOR(k, 1)
    
    Next k
Next i


'ATEMP_MATRIX = MATRIX_LU_LINEAR_SYSTEM_FUNC(BTEMP_MATRIX, XTEMP_VECTOR)
ATEMP_MATRIX = MATRIX_TRIANGULAR_LINEAR_SYSTEM_FUNC(BTEMP_MATRIX)

 'If IsArray(ATEMP_MATRIX) = True Then 'rebuild polinomial coefficients
If ATEMP_MATRIX(2) <> 0 Then  'load and mop-up numerator
    ReDim NUMER_VECTOR(0 To NUMER_DEG, 1 To 1)
    ReDim DENOM_VECTOR(0 To DENOM_DEG, 1 To 1)
    
'    PARAM_VECTOR = ATEMP_MATRIX
    PARAM_VECTOR = ATEMP_MATRIX(1)
    For i = 0 To NUMER_DEG
        NUMER_VECTOR(i, 1) = PARAM_VECTOR(i + 1, 1)
        If Abs(NUMER_VECTOR(i, 1)) < tolerance Then _
        NUMER_VECTOR(i, 1) = 0 'Else i0 = i
    Next i
    
    'load and mop-up denominator
    For i = 0 To DENOM_DEG - 1
        DENOM_VECTOR(i, 1) = PARAM_VECTOR(NUMER_DEG + i + 2, 1)
        If Abs(DENOM_VECTOR(i, 1)) < tolerance Then DENOM_VECTOR(i, 1) = 0
    Next i
    DENOM_VECTOR(DENOM_DEG, 1) = 1

Else
    GoTo ERROR_LABEL
End If

    'build the parameters vector
    'c = [a0, a1... an, b0, b1...bm-1]
    'remember that denominator polynomial is monic, i.e.  bm =1
ReDim PARAM_VECTOR(1 To (UBound(NUMER_VECTOR, 1) + 1) + _
                            (UBound(DENOM_VECTOR, 1) + 0), 1 To 1)
    
For i = 0 To UBound(NUMER_VECTOR, 1)
    PARAM_VECTOR(i + 1, 1) = NUMER_VECTOR(i, 1)
Next i
For i = 0 To UBound(DENOM_VECTOR, 1) - 1 'Exclude the last entry
        'of the Denom_Coef, which is = 1
    PARAM_VECTOR(i + UBound(NUMER_VECTOR, 1) + 2, 1) = DENOM_VECTOR(i, 1)
Next i

PUB_RATION_NUMER = NUMER_DEG
PUB_RATION_DENOM = DENOM_DEG

Select Case OUTPUT
'-----------------------------------------------------------------------------
Case 0
'-----------------------------------------------------------------------------
    XTEMP_VECTOR = LEVENBERG_MARQUARDT_OPTIMIZATION2_FUNC(XDATA_VECTOR, YDATA_VECTOR, _
                PARAM_VECTOR, "CALL_RATIONAL_OBJ_FUNC", _
                "CALL_RATIONAL_GRAD_FUNC", 0, nLOOPS, tolerance, epsilon)

    'reload polynomial coefficients
    j = 0
    For i = 0 To UBound(NUMER_VECTOR, 1)
        j = j + 1
        NUMER_VECTOR(i, 1) = XTEMP_VECTOR(j, 1)
    Next i
    For i = 0 To UBound(DENOM_VECTOR, 1) - 1
        j = j + 1
        DENOM_VECTOR(i, 1) = XTEMP_VECTOR(j, 1)
    Next i

    ReDim XTEMP_VECTOR(0 To (UBound(NUMER_VECTOR, 1) + 1) + _
                        (UBound(DENOM_VECTOR, 1) + 1), 1 To 2) As Variant

    XTEMP_VECTOR(0, 1) = CStr("FORMULA: ")
    XTEMP_VECTOR(0, 2) = RATIONAL_POLYNOMIAL_WRITE_FUNC(NUMER_DEG, DENOM_DEG, "a", "x")
    For i = 0 To UBound(NUMER_VECTOR, 1)
        XTEMP_VECTOR(i + 1, 1) = CStr("a" & i + 1 & " :")
        XTEMP_VECTOR(i + 1, 2) = NUMER_VECTOR(i, 1)
    Next i

    For i = 0 To UBound(DENOM_VECTOR, 1) 'Exclude the last entry
        'of the Denom_Coef, which is = 1
        XTEMP_VECTOR(i + UBound(NUMER_VECTOR, 1) + 2, 1) = _
        CStr("a" & i + UBound(NUMER_VECTOR, 1) + 2 & " :")
        XTEMP_VECTOR(i + UBound(NUMER_VECTOR, 1) + 2, 2) = DENOM_VECTOR(i, 1)
    Next i

    LEVENBERG_MARQUARDT_RATIONAL_FUNC = XTEMP_VECTOR
'-----------------------------------------------------------------------------
Case 1
'-----------------------------------------------------------------------------
    LEVENBERG_MARQUARDT_RATIONAL_FUNC = LEVENBERG_MARQUARDT_OPTIMIZATION2_FUNC(XDATA_VECTOR, YDATA_VECTOR, _
                PARAM_VECTOR, "CALL_RATIONAL_OBJ_FUNC", _
                "CALL_RATIONAL_GRAD_FUNC", 0, nLOOPS, tolerance, epsilon)
                
               
'-----------------------------------------------------------------------------
Case 2
'-----------------------------------------------------------------------------
    LEVENBERG_MARQUARDT_RATIONAL_FUNC = LEVENBERG_MARQUARDT_OPTIMIZATION2_FUNC(XDATA_VECTOR, YDATA_VECTOR, _
                PARAM_VECTOR, "CALL_RATIONAL_OBJ_FUNC", _
                "CALL_RATIONAL_GRAD_FUNC", 1, nLOOPS, tolerance, epsilon)
'-----------------------------------------------------------------------------
Case 3
'-----------------------------------------------------------------------------
    LEVENBERG_MARQUARDT_RATIONAL_FUNC = LEVENBERG_MARQUARDT_OPTIMIZATION2_FUNC(XDATA_VECTOR, YDATA_VECTOR, _
                PARAM_VECTOR, "CALL_RATIONAL_OBJ_FUNC", _
                "CALL_RATIONAL_GRAD_FUNC", 2, nLOOPS, tolerance, epsilon)
'-----------------------------------------------------------------------------
Case 4
'-----------------------------------------------------------------------------
    LEVENBERG_MARQUARDT_RATIONAL_FUNC = LEVENBERG_MARQUARDT_OPTIMIZATION2_FUNC(XDATA_VECTOR, YDATA_VECTOR, _
                PARAM_VECTOR, "CALL_RATIONAL_OBJ_FUNC", _
                "CALL_RATIONAL_GRAD_FUNC", 3, nLOOPS, tolerance, epsilon)
'-----------------------------------------------------------------------------
Case 5
'-----------------------------------------------------------------------------
    LEVENBERG_MARQUARDT_RATIONAL_FUNC = LEVENBERG_MARQUARDT_OPTIMIZATION2_FUNC(XDATA_VECTOR, YDATA_VECTOR, _
                PARAM_VECTOR, "CALL_RATIONAL_OBJ_FUNC", _
                "CALL_RATIONAL_GRAD_FUNC", 4, nLOOPS, tolerance, epsilon)
'-----------------------------------------------------------------------------
Case Else
'-----------------------------------------------------------------------------
    LEVENBERG_MARQUARDT_RATIONAL_FUNC = LEVENBERG_MARQUARDT_OPTIMIZATION2_FUNC(XDATA_VECTOR, YDATA_VECTOR, _
                PARAM_VECTOR, "CALL_RATIONAL_OBJ_FUNC", _
                "CALL_RATIONAL_GRAD_FUNC", 5, nLOOPS, tolerance, epsilon)
'-----------------------------------------------------------------------------
End Select
'-----------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
LEVENBERG_MARQUARDT_RATIONAL_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_RATIONAL_OBJ_FUNC
'DESCRIPTION   :
'LIBRARY       : OPTIMIZATION
'GROUP         : NLR_FRAMES
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function CALL_RATIONAL_OBJ_FUNC(ByRef XDATA_RNG As Variant, _
ByRef PARAM_RNG As Variant)

Dim XDATA_VECTOR As Variant
Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL
XDATA_VECTOR = XDATA_RNG
If UBound(XDATA_VECTOR, 1) = 1 Then: XDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(XDATA_VECTOR)
PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)

CALL_RATIONAL_OBJ_FUNC = RATIONAL_POLYNOMIAL_QUOTIENT_FUNC(XDATA_RNG, _
                         PARAM_RNG, PUB_RATION_NUMER, PUB_RATION_DENOM)
Exit Function
ERROR_LABEL:
CALL_RATIONAL_OBJ_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_RATIONAL_GRAD_FUNC
'DESCRIPTION   :
'LIBRARY       : OPTIMIZATION
'GROUP         : NLR_FRAMES
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function CALL_RATIONAL_GRAD_FUNC(ByRef XDATA_RNG As Variant, _
ByRef PARAM_RNG As Variant)

Dim XDATA_VECTOR As Variant
Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL
XDATA_VECTOR = XDATA_RNG
If UBound(XDATA_VECTOR, 1) = 1 Then: XDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(XDATA_VECTOR)
PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)

CALL_RATIONAL_GRAD_FUNC = RATIONAL_POLYNOMIAL_GRADIENT_FUNC(XDATA_RNG, _
                         PARAM_RNG, PUB_RATION_NUMER, PUB_RATION_DENOM)
Exit Function
ERROR_LABEL:
CALL_RATIONAL_GRAD_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : LEVENBERG_MARQUARDT_EXPONENTIAL_FUNC

'DESCRIPTION: Exponential fitting: 2-6 parameters
'Exponential relations are very common in the real world. Usually
'they appear together with oscillating circular function sine, cosine. This
'function is very useful to analyze a regression of a simple exponential and
'sum of exponentials.

'Many books suggest to transform this nonlinear function into a "linearized"
'one and then apply the linear regression to this new model. Performing the
'linear regression for this model we get the parameters (b0 , b1 ) and
'finally the original parameter of the nonlinear function (A , k) by these
'simple formulas. This method is quite popular but we have to put in evidence
'that this method could fail.

'In fact this is not a true "least squares nonlinear" regression, but a sort
'of quick method to obtain an approximation of the true "least squares
'nonlinear" regression. Sometime the parameters obtained by the linearization
'method are sufficiently close to those of the NL-LS (Non-Linear Least Squares)
'method; but sometime not and sometime could gives values completely different.
'So a good technique, always valid to check the result, is to calculate the
'residuals of the regression. If the least squares of the residual are too
'high the linearized regression must be rejected.

'Sometime, the parameters obtained by the linearized method could be used as
'starting point for the optimization algorithms performing the true NL-LS
'regression, like this function. The difference between linearized and true
'NL-LS regression exists for all models: logarithm, exponential, power, etc)
'but the difference may be so evident only for exponential functions. For
'other models the difference is always very low. There is another reason to
'dedicated many attention to this important but tricky regression

'LIBRARY       : OPTIMIZATION
'GROUP         : NLR_FRAMES
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************


Function LEVENBERG_MARQUARDT_EXPONENTIAL_FUNC(ByRef XDATA_RNG As Variant, _
ByRef YDATA_RNG As Variant, _
ByRef PARAM_RNG As Variant, _
Optional ByVal nLOOPS As Long = 1000, _
Optional ByVal tolerance As Double = 10 ^ -10, _
Optional ByVal epsilon As Double = 10 ^ -6, _
Optional ByVal OUTPUT As Integer = 0)

Dim NROWS As Long

Dim PARAM_VECTOR As Variant
Dim XDATA_MATRIX As Variant
Dim YDATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

XDATA_MATRIX = XDATA_RNG
If UBound(XDATA_MATRIX, 1) = 1 Then: XDATA_MATRIX = MATRIX_TRANSPOSE_FUNC(XDATA_MATRIX)

YDATA_VECTOR = YDATA_RNG
If UBound(YDATA_VECTOR, 1) = 1 Then: YDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(YDATA_VECTOR)

If UBound(XDATA_MATRIX, 1) <> UBound(YDATA_VECTOR, 1) Then: GoTo ERROR_LABEL
NROWS = UBound(XDATA_MATRIX, 1)

PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)
    
'--------------------------------------------------------------------------------
Select Case OUTPUT
'--------------------------------------------------------------------------------
Case 0
'--------------------------------------------------------------------------------
    LEVENBERG_MARQUARDT_EXPONENTIAL_FUNC = LEVENBERG_MARQUARDT_OPTIMIZATION2_FUNC(XDATA_MATRIX, YDATA_VECTOR, _
                PARAM_VECTOR, "CALL_EXPONENTIAL_OBJ_FUNC", _
                "CALL_EXPONENTIAL_GRAD_FUNC", 0, nLOOPS, tolerance, epsilon)
'--------------------------------------------------------------------------------
Case 1
'--------------------------------------------------------------------------------
    LEVENBERG_MARQUARDT_EXPONENTIAL_FUNC = LEVENBERG_MARQUARDT_OPTIMIZATION2_FUNC(XDATA_MATRIX, YDATA_VECTOR, _
                PARAM_VECTOR, "CALL_EXPONENTIAL_OBJ_FUNC", _
                "CALL_EXPONENTIAL_GRAD_FUNC", 1, nLOOPS, tolerance, epsilon)

'--------------------------------------------------------------------------------
Case 2
'--------------------------------------------------------------------------------
    LEVENBERG_MARQUARDT_EXPONENTIAL_FUNC = LEVENBERG_MARQUARDT_OPTIMIZATION2_FUNC(XDATA_MATRIX, YDATA_VECTOR, _
                PARAM_VECTOR, "CALL_EXPONENTIAL_OBJ_FUNC", _
                "CALL_EXPONENTIAL_GRAD_FUNC", 2, nLOOPS, tolerance, epsilon)

'--------------------------------------------------------------------------------
Case 3
'--------------------------------------------------------------------------------
    LEVENBERG_MARQUARDT_EXPONENTIAL_FUNC = LEVENBERG_MARQUARDT_OPTIMIZATION2_FUNC(XDATA_MATRIX, YDATA_VECTOR, _
                PARAM_VECTOR, "CALL_EXPONENTIAL_OBJ_FUNC", _
                "CALL_EXPONENTIAL_GRAD_FUNC", 3, nLOOPS, tolerance, epsilon)

'--------------------------------------------------------------------------------
Case 4
'--------------------------------------------------------------------------------
    LEVENBERG_MARQUARDT_EXPONENTIAL_FUNC = LEVENBERG_MARQUARDT_OPTIMIZATION2_FUNC(XDATA_MATRIX, YDATA_VECTOR, _
                PARAM_VECTOR, "CALL_EXPONENTIAL_OBJ_FUNC", _
                "CALL_EXPONENTIAL_GRAD_FUNC", 4, nLOOPS, tolerance, epsilon)

'--------------------------------------------------------------------------------
Case Else
'--------------------------------------------------------------------------------
    LEVENBERG_MARQUARDT_EXPONENTIAL_FUNC = LEVENBERG_MARQUARDT_OPTIMIZATION2_FUNC(XDATA_MATRIX, YDATA_VECTOR, _
                PARAM_VECTOR, "CALL_EXPONENTIAL_OBJ_FUNC", _
                "CALL_EXPONENTIAL_GRAD_FUNC", 5, nLOOPS, tolerance, epsilon)

'--------------------------------------------------------------------------------
End Select
'--------------------------------------------------------------------------------


Exit Function
ERROR_LABEL:
LEVENBERG_MARQUARDT_EXPONENTIAL_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_EXPONENTIAL_OBJ_FUNC
'DESCRIPTION   :
'LIBRARY       : OPTIMIZATION
'GROUP         : NLR_FRAMES
'ID            : 007
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function CALL_EXPONENTIAL_OBJ_FUNC(ByRef XDATA_RNG As Variant, _
ByRef PARAM_RNG As Variant)

Dim i As Long
Dim NROWS As Long
Dim NSIZE As Long
Dim PARAM_VECTOR As Variant
Dim XDATA_MATRIX As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)
NSIZE = UBound(PARAM_VECTOR, 1)

XDATA_MATRIX = XDATA_RNG
If UBound(XDATA_MATRIX, 1) = 1 Then: XDATA_MATRIX = MATRIX_TRANSPOSE_FUNC(XDATA_MATRIX)

NROWS = UBound(XDATA_MATRIX, 1)

ReDim TEMP_MATRIX(1 To NROWS, 1 To 1)

'-----------------------------------------------------------------------------------
Select Case NSIZE
'-----------------------------------------------------------------------------------
Case 2
'-----------------------------------------------------------------------------------
    For i = 1 To NROWS
        TEMP_MATRIX(i, 1) = PARAM_VECTOR(1, 1) * _
                    Exp(PARAM_VECTOR(2, 1) * XDATA_MATRIX(i, 1))
    Next i
'-----------------------------------------------------------------------------------
Case 3
'-----------------------------------------------------------------------------------
    For i = 1 To NROWS
        TEMP_MATRIX(i, 1) = PARAM_VECTOR(1, 1) + _
                    PARAM_VECTOR(2, 1) * Exp(PARAM_VECTOR(3, 1) * _
                            XDATA_MATRIX(i, 1))
    Next i
'-----------------------------------------------------------------------------------
Case 4
'-----------------------------------------------------------------------------------
    For i = 1 To NROWS
            TEMP_MATRIX(i, 1) = PARAM_VECTOR(1, 1) * _
                    Exp(PARAM_VECTOR(2, 1) * XDATA_MATRIX(i, 1)) + _
                            PARAM_VECTOR(3, 1) * Exp(PARAM_VECTOR(4, 1) * _
                                    XDATA_MATRIX(i, 1))
    Next i

'-----------------------------------------------------------------------------------
Case 5
'-----------------------------------------------------------------------------------
    For i = 1 To NROWS
            TEMP_MATRIX(i, 1) = PARAM_VECTOR(1, 1) + PARAM_VECTOR(2, 1) * _
                        Exp(PARAM_VECTOR(3, 1) * XDATA_MATRIX(i, 1)) + _
                                PARAM_VECTOR(4, 1) * Exp(PARAM_VECTOR(5, 1) * _
                                        XDATA_MATRIX(i, 1))
    Next i

'-----------------------------------------------------------------------------------
Case 6
    For i = 1 To NROWS
'-----------------------------------------------------------------------------------
            TEMP_MATRIX(i, 1) = PARAM_VECTOR(1, 1) * Exp(PARAM_VECTOR(2, 1) * _
                        XDATA_MATRIX(i, 1)) + PARAM_VECTOR(3, 1) * _
                                Exp(PARAM_VECTOR(4, 1) * XDATA_MATRIX(i, 1)) + _
                                        PARAM_VECTOR(5, 1) * _
                                                Exp(PARAM_VECTOR(6, 1) * _
                                                        XDATA_MATRIX(i, 1))
    Next i
'-----------------------------------------------------------------------------------
Case Else
'-----------------------------------------------------------------------------------
    GoTo ERROR_LABEL
'-----------------------------------------------------------------------------------
End Select
'-----------------------------------------------------------------------------------


CALL_EXPONENTIAL_OBJ_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
CALL_EXPONENTIAL_OBJ_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_EXPONENTIAL_GRAD_FUNC
'DESCRIPTION   :
'LIBRARY       : OPTIMIZATION
'GROUP         : NLR_FRAMES
'ID            : 008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function CALL_EXPONENTIAL_GRAD_FUNC(ByRef XDATA_RNG As Variant, _
ByRef PARAM_RNG As Variant)

Dim i As Long
Dim NROWS As Long
Dim NSIZE As Long
Dim PARAM_VECTOR As Variant
Dim XDATA_MATRIX As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)
NSIZE = UBound(PARAM_VECTOR, 1)

XDATA_MATRIX = XDATA_RNG
If UBound(XDATA_MATRIX, 1) = 1 Then: XDATA_MATRIX = MATRIX_TRANSPOSE_FUNC(XDATA_MATRIX)

NROWS = UBound(XDATA_MATRIX, 1)

ReDim TEMP_MATRIX(1 To NROWS, 1 To NSIZE)

'-----------------------------------------------------------------------------------
Select Case NSIZE
'-----------------------------------------------------------------------------------
Case 2
'-----------------------------------------------------------------------------------
    For i = 1 To NROWS
        TEMP_MATRIX(i, 1) = Exp(PARAM_VECTOR(2, 1) * XDATA_MATRIX(i, 1))
        TEMP_MATRIX(i, 2) = XDATA_MATRIX(i, 1) * PARAM_VECTOR(1, 1) * _
                            Exp(PARAM_VECTOR(2, 1) * XDATA_MATRIX(i, 1))
    Next i
'-----------------------------------------------------------------------------------
Case 3
'-----------------------------------------------------------------------------------
    For i = 1 To NROWS
        TEMP_MATRIX(i, 1) = 1
        TEMP_MATRIX(i, 2) = Exp(PARAM_VECTOR(3, 1) * XDATA_MATRIX(i, 1))
        TEMP_MATRIX(i, 3) = XDATA_MATRIX(i, 1) * PARAM_VECTOR(2, 1) * _
                        Exp(PARAM_VECTOR(3, 1) * XDATA_MATRIX(i, 1))
    Next i

'-----------------------------------------------------------------------------------
Case 4
'-----------------------------------------------------------------------------------
    For i = 1 To NROWS
        TEMP_MATRIX(i, 1) = Exp(PARAM_VECTOR(2, 1) * XDATA_MATRIX(i, 1))
        TEMP_MATRIX(i, 2) = XDATA_MATRIX(i, 1) * PARAM_VECTOR(1, 1) * _
                        Exp(PARAM_VECTOR(2, 1) * XDATA_MATRIX(i, 1))
        TEMP_MATRIX(i, 3) = Exp(PARAM_VECTOR(4, 1) * XDATA_MATRIX(i, 1))
        TEMP_MATRIX(i, 4) = XDATA_MATRIX(i, 1) * PARAM_VECTOR(3, 1) * _
                    Exp(PARAM_VECTOR(4, 1) * XDATA_MATRIX(i, 1))
    Next i
'-----------------------------------------------------------------------------------
Case 5
'-----------------------------------------------------------------------------------
    For i = 1 To NROWS
        TEMP_MATRIX(i, 1) = 1
        TEMP_MATRIX(i, 2) = Exp(PARAM_VECTOR(3, 1) * XDATA_MATRIX(i, 1))
        TEMP_MATRIX(i, 3) = PARAM_VECTOR(2, 1) * XDATA_MATRIX(i, 1) * _
                    Exp(PARAM_VECTOR(3, 1) * XDATA_MATRIX(i, 1))
        TEMP_MATRIX(i, 4) = Exp(PARAM_VECTOR(5, 1) * XDATA_MATRIX(i, 1))
        TEMP_MATRIX(i, 5) = PARAM_VECTOR(4, 1) * XDATA_MATRIX(i, 1) * _
                    Exp(PARAM_VECTOR(5, 1) * XDATA_MATRIX(i, 1))
    Next i

'-----------------------------------------------------------------------------------
Case 6
'-----------------------------------------------------------------------------------
    For i = 1 To NROWS
        TEMP_MATRIX(i, 1) = Exp(PARAM_VECTOR(2, 1) * XDATA_MATRIX(i, 1))
        TEMP_MATRIX(i, 2) = XDATA_MATRIX(i, 1) * PARAM_VECTOR(1, 1) * _
                    Exp(PARAM_VECTOR(2, 1) * XDATA_MATRIX(i, 1))
        TEMP_MATRIX(i, 3) = Exp(PARAM_VECTOR(4, 1) * XDATA_MATRIX(i, 1))
        TEMP_MATRIX(i, 4) = XDATA_MATRIX(i, 1) * PARAM_VECTOR(3, 1) * _
                    Exp(PARAM_VECTOR(4, 1) * XDATA_MATRIX(i, 1))
        TEMP_MATRIX(i, 5) = Exp(PARAM_VECTOR(6, 1) * XDATA_MATRIX(i, 1))
        TEMP_MATRIX(i, 6) = XDATA_MATRIX(i, 1) * PARAM_VECTOR(5, 1) * _
                    Exp(PARAM_VECTOR(6, 1) * XDATA_MATRIX(i, 1))
    Next i
'-----------------------------------------------------------------------------------
Case Else
'-----------------------------------------------------------------------------------
    GoTo ERROR_LABEL
'-----------------------------------------------------------------------------------
End Select
'-----------------------------------------------------------------------------------

CALL_EXPONENTIAL_GRAD_FUNC = TEMP_MATRIX
Exit Function
ERROR_LABEL:
CALL_EXPONENTIAL_GRAD_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : LEVENBERG_MARQUARDT_DAMPED_COSINE_FUNC

'DESCRIPTION: Damped cosine regression: That is a very common behaviour of
'a 2nd order real system. The responce oscillates around a final value with
'amplitude decreasing with the time --> y = a0 + a1 * e^(kt) * cos(w*t+0)

'This model has 5 parameters:
'a0 offset or final value
'a1 amplitude.
'k damping factor.
'w pulsation (2*pi*f )
'0 phase

'Related to this model is the following one, called "damped-sine-cosine"
'y = a0 + e^(k*x) * (bc*Cos(w*t) + bs*Sin(w*t))
'where : bc = a1 * Cos(0); bs = -a1 * Sin(0)

'LIBRARY       : OPTIMIZATION
'GROUP         : NLR_FRAMES
'ID            : 009
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function LEVENBERG_MARQUARDT_DAMPED_COSINE_FUNC(ByRef XDATA_RNG As Variant, _
ByRef YDATA_RNG As Variant, _
ByRef PARAM_RNG As Variant, _
Optional ByVal nLOOPS As Long = 1000, _
Optional ByVal tolerance As Double = 10 ^ -10, _
Optional ByVal epsilon As Double = 10 ^ -6, _
Optional ByVal OUTPUT As Integer = 0)

'Dumped Exponential cosine

Dim NROWS As Long

Dim PARAM_VECTOR As Variant
Dim XDATA_MATRIX As Variant
Dim YDATA_VECTOR As Variant

On Error GoTo ERROR_LABEL


XDATA_MATRIX = XDATA_RNG
If UBound(XDATA_MATRIX, 1) = 1 Then: XDATA_MATRIX = MATRIX_TRANSPOSE_FUNC(XDATA_MATRIX)

YDATA_VECTOR = YDATA_RNG
If UBound(YDATA_VECTOR, 1) = 1 Then: YDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(YDATA_VECTOR)

If UBound(XDATA_MATRIX, 1) <> UBound(YDATA_VECTOR, 1) Then: GoTo ERROR_LABEL
NROWS = UBound(XDATA_MATRIX, 1)


PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)
If UBound(PARAM_VECTOR, 1) <> 5 Then: GoTo ERROR_LABEL
'parameters missing or too many parameters


'--------------------------------------------------------------------------------
Select Case OUTPUT
'--------------------------------------------------------------------------------
Case 0
'--------------------------------------------------------------------------------
'PARAM_VECTOR:
'   ROW(1) --> offset
'   ROW(2) --> amp.
'   ROW(3) --> damp.
'   ROW(4) --> puls.
'   ROW(5) --> phase

    LEVENBERG_MARQUARDT_DAMPED_COSINE_FUNC = LEVENBERG_MARQUARDT_OPTIMIZATION2_FUNC(XDATA_MATRIX, YDATA_VECTOR, _
                PARAM_VECTOR, "CALL_DAMPED_COSINE_OBJ_FUNC", _
                "CALL_DAMPED_COSINE_GRAD_FUNC", 0, nLOOPS, tolerance, epsilon)
'--------------------------------------------------------------------------------
Case 1
'--------------------------------------------------------------------------------
    LEVENBERG_MARQUARDT_DAMPED_COSINE_FUNC = LEVENBERG_MARQUARDT_OPTIMIZATION2_FUNC(XDATA_MATRIX, YDATA_VECTOR, _
                PARAM_VECTOR, "CALL_DAMPED_COSINE_OBJ_FUNC", _
                "CALL_DAMPED_COSINE_GRAD_FUNC", 1, nLOOPS, tolerance, epsilon)

'--------------------------------------------------------------------------------
Case 2
'--------------------------------------------------------------------------------
    LEVENBERG_MARQUARDT_DAMPED_COSINE_FUNC = LEVENBERG_MARQUARDT_OPTIMIZATION2_FUNC(XDATA_MATRIX, YDATA_VECTOR, _
                PARAM_VECTOR, "CALL_DAMPED_COSINE_OBJ_FUNC", _
                "CALL_DAMPED_COSINE_GRAD_FUNC", 2, nLOOPS, tolerance, epsilon)

'--------------------------------------------------------------------------------
Case 3
'--------------------------------------------------------------------------------
    LEVENBERG_MARQUARDT_DAMPED_COSINE_FUNC = LEVENBERG_MARQUARDT_OPTIMIZATION2_FUNC(XDATA_MATRIX, YDATA_VECTOR, _
                PARAM_VECTOR, "CALL_DAMPED_COSINE_OBJ_FUNC", _
                "CALL_DAMPED_COSINE_GRAD_FUNC", 3, nLOOPS, tolerance, epsilon)

'--------------------------------------------------------------------------------
Case 4
'--------------------------------------------------------------------------------
    LEVENBERG_MARQUARDT_DAMPED_COSINE_FUNC = LEVENBERG_MARQUARDT_OPTIMIZATION2_FUNC(XDATA_MATRIX, YDATA_VECTOR, _
                PARAM_VECTOR, "CALL_DAMPED_COSINE_OBJ_FUNC", _
                "CALL_DAMPED_COSINE_GRAD_FUNC", 4, nLOOPS, tolerance, epsilon)

'--------------------------------------------------------------------------------
Case Else
'--------------------------------------------------------------------------------
    LEVENBERG_MARQUARDT_DAMPED_COSINE_FUNC = LEVENBERG_MARQUARDT_OPTIMIZATION2_FUNC(XDATA_MATRIX, YDATA_VECTOR, _
                PARAM_VECTOR, "CALL_DAMPED_COSINE_OBJ_FUNC", _
                "CALL_DAMPED_COSINE_GRAD_FUNC", 5, nLOOPS, tolerance, epsilon)

'--------------------------------------------------------------------------------
End Select
'--------------------------------------------------------------------------------


Exit Function
ERROR_LABEL:
LEVENBERG_MARQUARDT_DAMPED_COSINE_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_DAMPED_COSINE_OBJ_FUNC
'DESCRIPTION   :
'LIBRARY       : OPTIMIZATION
'GROUP         : NLR_FRAMES
'ID            : 010
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function CALL_DAMPED_COSINE_OBJ_FUNC(ByRef XDATA_RNG As Variant, _
ByRef PARAM_RNG As Variant)

Dim i As Long
Dim NROWS As Long
Dim NSIZE As Long
Dim PARAM_VECTOR As Variant
Dim XDATA_MATRIX As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)
NSIZE = UBound(PARAM_VECTOR, 1)
If NSIZE <> 5 Then: GoTo ERROR_LABEL

XDATA_MATRIX = XDATA_RNG
If UBound(XDATA_MATRIX, 1) = 1 Then: XDATA_MATRIX = MATRIX_TRANSPOSE_FUNC(XDATA_MATRIX)

NROWS = UBound(XDATA_MATRIX, 1)

ReDim TEMP_MATRIX(1 To NROWS, 1 To 1)

For i = 1 To NROWS
    TEMP_MATRIX(i, 1) = PARAM_VECTOR(1, 1) + PARAM_VECTOR(2, 1) * _
                            Exp(PARAM_VECTOR(3, 1) * XDATA_MATRIX(i, 1)) * _
                                        Cos(PARAM_VECTOR(4, 1) * _
                                            XDATA_MATRIX(i, 1) + _
                                            PARAM_VECTOR(5, 1))
Next i

CALL_DAMPED_COSINE_OBJ_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
CALL_DAMPED_COSINE_OBJ_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_DAMPED_COSINE_GRAD_FUNC
'DESCRIPTION   :
'LIBRARY       : OPTIMIZATION
'GROUP         : NLR_FRAMES
'ID            : 011
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function CALL_DAMPED_COSINE_GRAD_FUNC(ByRef XDATA_RNG As Variant, _
ByRef PARAM_RNG As Variant)

Dim i As Long
Dim NROWS As Long
Dim NSIZE As Long
Dim PARAM_VECTOR As Variant
Dim XDATA_MATRIX As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)
NSIZE = UBound(PARAM_VECTOR, 1)
If NSIZE <> 5 Then: GoTo ERROR_LABEL

XDATA_MATRIX = XDATA_RNG
If UBound(XDATA_MATRIX, 1) = 1 Then: XDATA_MATRIX = MATRIX_TRANSPOSE_FUNC(XDATA_MATRIX)

NROWS = UBound(XDATA_MATRIX, 1)

ReDim TEMP_MATRIX(1 To NROWS, 1 To 5)

For i = 1 To NROWS
    TEMP_MATRIX(i, 1) = 1
    TEMP_MATRIX(i, 2) = Exp(PARAM_VECTOR(3, 1) * XDATA_MATRIX(i, 1)) * _
                            Cos(PARAM_VECTOR(4, 1) * XDATA_MATRIX(i, 1) + _
                                PARAM_VECTOR(5, 1))
    TEMP_MATRIX(i, 3) = XDATA_MATRIX(i, 1) * PARAM_VECTOR(2, 1) * _
                            TEMP_MATRIX(i, 2)
    TEMP_MATRIX(i, 4) = -XDATA_MATRIX(i, 1) * PARAM_VECTOR(2, 1) * _
                            Exp(PARAM_VECTOR(3, 1) * XDATA_MATRIX(i, 1)) * _
                                Sin(PARAM_VECTOR(4, 1) * XDATA_MATRIX(i, 1) + _
                                    PARAM_VECTOR(5, 1))
    TEMP_MATRIX(i, 5) = -PARAM_VECTOR(2, 1) * Exp(PARAM_VECTOR(3, 1) * _
                            XDATA_MATRIX(i, 1)) * Sin(PARAM_VECTOR(4, 1) * _
                                XDATA_MATRIX(i, 1) + PARAM_VECTOR(5, 1))
Next i

CALL_DAMPED_COSINE_GRAD_FUNC = TEMP_MATRIX
Exit Function
ERROR_LABEL:
CALL_DAMPED_COSINE_GRAD_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : LEVENBERG_MARQUARDT_GAUSS_FUNC

'DESCRIPTION:
'DESCRIPTION   : Gaussian regression is a symmetrical exponential model
'useful for many applications. Three parameters determine completely the
'Gaussian function a = amplitude; b = axis of symmetry; c = deviation Or spread
'By inspection of the plot, it is quite easy to recognize and evaluate
'these parameters: The axis of symmetry "b" is the abscissa where the
'function has its maximum amplitude of "a". The deviation "c" is the
'length where the function is at 37 % of its maximum value.

'Usually the data are affected by several errors that "mask" the original
'Gaussian distribution. In that case we can use the regression method to
'measure how the row data fit the Gaussian model and to evaluate its parameters.

'LIBRARY       : OPTIMIZATION
'GROUP         : NLR_FRAMES
'ID            : 012
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function LEVENBERG_MARQUARDT_GAUSS_FUNC(ByRef XDATA_RNG As Variant, _
ByRef YDATA_RNG As Variant, _
Optional ByRef PARAM_RNG As Variant, _
Optional ByVal nLOOPS As Long = 400, _
Optional ByVal tolerance As Double = 2 * 10 ^ -16, _
Optional ByVal epsilon As Double = 10 ^ -8, _
Optional ByVal OUTPUT As Integer = 0)

Dim i As Long
Dim NROWS As Long

Dim SY_VAL As Double
Dim SXY_VAL As Double
Dim YUP_VAL As Double
Dim TEMP_SUM As Double

Dim PARAM_VECTOR As Variant
Dim XDATA_MATRIX As Variant
Dim YDATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

XDATA_MATRIX = XDATA_RNG
If UBound(XDATA_MATRIX, 1) = 1 Then: XDATA_MATRIX = MATRIX_TRANSPOSE_FUNC(XDATA_MATRIX)

YDATA_VECTOR = YDATA_RNG
If UBound(YDATA_VECTOR, 1) = 1 Then: YDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(YDATA_VECTOR)

If UBound(XDATA_MATRIX, 1) <> UBound(YDATA_VECTOR, 1) Then: GoTo ERROR_LABEL
NROWS = UBound(XDATA_MATRIX, 1)

If IsArray(PARAM_RNG) = True Then
    PARAM_VECTOR = PARAM_RNG
    If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)
    If UBound(PARAM_VECTOR, 1) <> 3 Then: GoTo ERROR_LABEL 'parameters mismatching
Else
    ReDim PARAM_VECTOR(1 To 3, 1 To 1) 'parameter estimation
    PARAM_VECTOR(1, 1) = 0
    SXY_VAL = 0
    SY_VAL = 0
    For i = 2 To NROWS - 1
        SXY_VAL = SXY_VAL + XDATA_MATRIX(i, 1) * YDATA_VECTOR(i, 1)
        SY_VAL = SY_VAL + YDATA_VECTOR(i, 1)
        YUP_VAL = (YDATA_VECTOR(i - 1, 1) + YDATA_VECTOR(i, 1) + _
                   YDATA_VECTOR(i + 1, 1)) / 3
        If YUP_VAL > PARAM_VECTOR(1, 1) Then: PARAM_VECTOR(1, 1) = YUP_VAL
    Next i
    If SY_VAL = 0 Then: GoTo ERROR_LABEL 'Unable to found regression parameters
    PARAM_VECTOR(2, 1) = SXY_VAL / SY_VAL
    TEMP_SUM = 0
    For i = 2 To NROWS - 1
        TEMP_SUM = TEMP_SUM + YDATA_VECTOR(i, 1) * _
                    Abs(XDATA_MATRIX(i, 1) - PARAM_VECTOR(2, 1))
    Next i
    PARAM_VECTOR(3, 1) = Sqr(TEMP_SUM / SY_VAL)
End If

'--------------------------------------------------------------------------------
Select Case OUTPUT
'--------------------------------------------------------------------------------
Case 0
'--------------------------------------------------------------------------------
    
    LEVENBERG_MARQUARDT_GAUSS_FUNC = LEVENBERG_MARQUARDT_OPTIMIZATION2_FUNC(XDATA_MATRIX, YDATA_VECTOR, _
                PARAM_VECTOR, "CALL_GAUSS_OBJ_FUNC", _
                "CALL_GAUSS_GRAD_FUNC", 0, nLOOPS)
'--------------------------------------------------------------------------------
Case 1
'--------------------------------------------------------------------------------
    LEVENBERG_MARQUARDT_GAUSS_FUNC = LEVENBERG_MARQUARDT_OPTIMIZATION2_FUNC(XDATA_MATRIX, YDATA_VECTOR, _
                PARAM_VECTOR, "CALL_GAUSS_OBJ_FUNC", _
                "CALL_GAUSS_GRAD_FUNC", 1, nLOOPS, tolerance, epsilon)

'--------------------------------------------------------------------------------
Case 2
'--------------------------------------------------------------------------------
    LEVENBERG_MARQUARDT_GAUSS_FUNC = LEVENBERG_MARQUARDT_OPTIMIZATION2_FUNC(XDATA_MATRIX, YDATA_VECTOR, _
                PARAM_VECTOR, "CALL_GAUSS_OBJ_FUNC", _
                "CALL_GAUSS_GRAD_FUNC", 2, nLOOPS, tolerance, epsilon)

'--------------------------------------------------------------------------------
Case 3
'--------------------------------------------------------------------------------
    LEVENBERG_MARQUARDT_GAUSS_FUNC = LEVENBERG_MARQUARDT_OPTIMIZATION2_FUNC(XDATA_MATRIX, YDATA_VECTOR, _
                PARAM_VECTOR, "CALL_GAUSS_OBJ_FUNC", _
                "CALL_GAUSS_GRAD_FUNC", 3, nLOOPS, tolerance, epsilon)

'--------------------------------------------------------------------------------
Case 4
'--------------------------------------------------------------------------------
    LEVENBERG_MARQUARDT_GAUSS_FUNC = LEVENBERG_MARQUARDT_OPTIMIZATION2_FUNC(XDATA_MATRIX, YDATA_VECTOR, _
                PARAM_VECTOR, "CALL_GAUSS_OBJ_FUNC", _
                "CALL_GAUSS_GRAD_FUNC", 4, nLOOPS, tolerance, epsilon)

'--------------------------------------------------------------------------------
Case Else
'--------------------------------------------------------------------------------
    LEVENBERG_MARQUARDT_GAUSS_FUNC = LEVENBERG_MARQUARDT_OPTIMIZATION2_FUNC(XDATA_MATRIX, YDATA_VECTOR, _
                PARAM_VECTOR, "CALL_GAUSS_OBJ_FUNC", _
                "CALL_GAUSS_GRAD_FUNC", 5, nLOOPS, tolerance, epsilon)

'--------------------------------------------------------------------------------
End Select
'--------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
LEVENBERG_MARQUARDT_GAUSS_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_GAUSS_OBJ_FUNC
'DESCRIPTION   :
'LIBRARY       : OPTIMIZATION
'GROUP         : NLR_FRAMES
'ID            : 013
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function CALL_GAUSS_OBJ_FUNC(ByRef XDATA_RNG As Variant, _
ByRef PARAM_RNG As Variant)

Dim i As Long
Dim NROWS As Long
Dim NSIZE As Long
Dim PARAM_VECTOR As Variant
Dim XDATA_MATRIX As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)
NSIZE = UBound(PARAM_VECTOR, 1)
If NSIZE <> 3 Then: GoTo ERROR_LABEL

XDATA_MATRIX = XDATA_RNG
If UBound(XDATA_MATRIX, 1) = 1 Then: XDATA_MATRIX = MATRIX_TRANSPOSE_FUNC(XDATA_MATRIX)

NROWS = UBound(XDATA_MATRIX, 1)

ReDim TEMP_MATRIX(1 To NROWS, 1 To 1)

For i = 1 To NROWS
    TEMP_MATRIX(i, 1) = PARAM_VECTOR(1, 1) * Exp(-(((XDATA_MATRIX(i, 1) - _
                         PARAM_VECTOR(2, 1)) / PARAM_VECTOR(3, 1)) ^ 2))
Next i


CALL_GAUSS_OBJ_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
CALL_GAUSS_OBJ_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_GAUSS_GRAD_FUNC
'DESCRIPTION   :
'LIBRARY       : OPTIMIZATION
'GROUP         : NLR_FRAMES
'ID            : 014
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function CALL_GAUSS_GRAD_FUNC(ByRef XDATA_RNG As Variant, _
ByRef PARAM_RNG As Variant)

Dim i As Long
Dim NROWS As Long
Dim NSIZE As Long
Dim PARAM_VECTOR As Variant
Dim XDATA_MATRIX As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)
NSIZE = UBound(PARAM_VECTOR, 1)
If NSIZE <> 3 Then: GoTo ERROR_LABEL

XDATA_MATRIX = XDATA_RNG
If UBound(XDATA_MATRIX, 1) = 1 Then: XDATA_MATRIX = MATRIX_TRANSPOSE_FUNC(XDATA_MATRIX)

NROWS = UBound(XDATA_MATRIX, 1)

ReDim TEMP_MATRIX(1 To NROWS, 1 To 3)

For i = 1 To NROWS
    TEMP_MATRIX(i, 1) = Exp(-(((XDATA_MATRIX(i, 1) - PARAM_VECTOR(2, 1)) / _
                            PARAM_VECTOR(3, 1)) ^ 2))
    TEMP_MATRIX(i, 2) = 2 * PARAM_VECTOR(1, 1) * (XDATA_MATRIX(i, 1) - _
                        PARAM_VECTOR(2, 1)) / PARAM_VECTOR(3, 1) ^ 2 * _
                            TEMP_MATRIX(i, 1)
    TEMP_MATRIX(i, 3) = (XDATA_MATRIX(i, 1) - PARAM_VECTOR(2, 1)) / _
                        PARAM_VECTOR(3, 1) * TEMP_MATRIX(i, 2)
Next i

CALL_GAUSS_GRAD_FUNC = TEMP_MATRIX
Exit Function
ERROR_LABEL:
CALL_GAUSS_GRAD_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : LEVENBERG_MARQUARDT_POWER_FUNC

'DESCRIPTION   : POWER regression
'The simple model of this regression is: y = a * x^k
'with x = 0 and k >0, a >0. When k < 0 the model becames
'y = a / x ^ -k; with x > 0 and a > 0 (x must be strictly positive)
'Usually this function is used to estimate the growth-rate or decay-rate
'of a population. We distingue the followinh cases of the exponent
'parameter k:

'1) In the case k > 1 the power fitting is very closed to
'the polynomial fitting when the exponent is positive and greater
'than 1 ( k > 1)

'2) In the case 0 < k < 1, when the exponent k is positive and
'lower then 1 , the polynomial model is unsuitable and the power model
'should be used. The difficult of this polynomial regression is located
'near the origin where the derivative grows sharply, becaming infinite
'for x = 0. Any polynomials, having no singular point, cannot follow the
'curve near this point. If we would have a dataset far from the singular
'point x = 0 , also polynomial regression would be better.

'3) In the case k < 0, when the exponent k is negative the power fitting
'should be used. Also a logarithmic regression could be used.

'LIBRARY       : OPTIMIZATION
'GROUP         : NLR_FRAMES
'ID            : 015
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function LEVENBERG_MARQUARDT_POWER_FUNC(ByRef XDATA_RNG As Variant, _
ByRef YDATA_RNG As Variant, _
Optional ByRef PARAM_RNG As Variant, _
Optional ByVal nLOOPS As Long = 1000, _
Optional ByVal tolerance As Double = 10 ^ -10, _
Optional ByVal epsilon As Double = 10 ^ -6, _
Optional ByVal OUTPUT As Integer = 0)
'power fitting

Dim i As Long
Dim NROWS As Long

Dim XDATA_MATRIX As Variant
Dim YDATA_VECTOR As Variant

Dim PARAM_VECTOR As Variant

Dim BTEMP_VECTOR As Variant

On Error GoTo ERROR_LABEL

XDATA_MATRIX = XDATA_RNG
If UBound(XDATA_MATRIX, 1) = 1 Then: XDATA_MATRIX = MATRIX_TRANSPOSE_FUNC(XDATA_MATRIX)

YDATA_VECTOR = YDATA_RNG
If UBound(YDATA_VECTOR, 1) = 1 Then: YDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(YDATA_VECTOR)

If UBound(XDATA_MATRIX, 1) <> UBound(YDATA_VECTOR, 1) Then: GoTo ERROR_LABEL
NROWS = UBound(XDATA_MATRIX, 1)

For i = 1 To NROWS
    If (XDATA_MATRIX(i, 1) <= 0) Or _
           (YDATA_VECTOR(i, 1) <= 0) Then
           GoTo ERROR_LABEL
           'negative values in the datasets
    End If
Next i

If IsArray(PARAM_RNG) = False Then 'parameter estimation
    
    ReDim ATEMP_VECTOR(1 To NROWS, 1 To 2)
    ReDim BTEMP_VECTOR(1 To NROWS, 1 To 1)

    For i = 1 To NROWS
        ATEMP_VECTOR(i, 1) = 1
        ATEMP_VECTOR(i, 2) = Log(XDATA_MATRIX(i, 1))
        BTEMP_VECTOR(i, 1) = Log(YDATA_VECTOR(i, 1))
    Next i
            
    PARAM_VECTOR = MATRIX_LEAST_SQUARE_LINEAR_SYSTEM_FUNC(ATEMP_VECTOR, BTEMP_VECTOR, 0)


    PARAM_VECTOR = MATRIX_TRIANGULAR_LINEAR_SYSTEM_FUNC(PARAM_VECTOR)
    If PARAM_VECTOR(UBound(PARAM_VECTOR)) <> 0 Then
        PARAM_VECTOR = PARAM_VECTOR(LBound(PARAM_VECTOR))
    Else
        GoTo ERROR_LABEL
    End If
    
    PARAM_VECTOR(1, 1) = Exp(PARAM_VECTOR(1, 1))
Else
    PARAM_VECTOR = PARAM_RNG
    If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)
    If UBound(PARAM_VECTOR, 1) <> 2 Then: GoTo ERROR_LABEL
End If

'PARAM_VECTOR --> [Amp,Exp]

'--------------------------------------------------------------------------------
Select Case OUTPUT
'--------------------------------------------------------------------------------
Case 0
'--------------------------------------------------------------------------------
    LEVENBERG_MARQUARDT_POWER_FUNC = LEVENBERG_MARQUARDT_OPTIMIZATION2_FUNC(XDATA_MATRIX, YDATA_VECTOR, _
                PARAM_VECTOR, "CALL_POWER_OBJ_FUNC", _
                "CALL_POWER_GRAD_FUNC", 0, nLOOPS, tolerance, epsilon)
'--------------------------------------------------------------------------------
Case 1
'--------------------------------------------------------------------------------
    LEVENBERG_MARQUARDT_POWER_FUNC = LEVENBERG_MARQUARDT_OPTIMIZATION2_FUNC(XDATA_MATRIX, YDATA_VECTOR, _
                PARAM_VECTOR, "CALL_POWER_OBJ_FUNC", _
                "CALL_POWER_GRAD_FUNC", 1, nLOOPS, tolerance, epsilon)

'--------------------------------------------------------------------------------
Case 2
'--------------------------------------------------------------------------------
    LEVENBERG_MARQUARDT_POWER_FUNC = LEVENBERG_MARQUARDT_OPTIMIZATION2_FUNC(XDATA_MATRIX, YDATA_VECTOR, _
                PARAM_VECTOR, "CALL_POWER_OBJ_FUNC", _
                "CALL_POWER_GRAD_FUNC", 2, nLOOPS, tolerance, epsilon)

'--------------------------------------------------------------------------------
Case 3
'--------------------------------------------------------------------------------
    LEVENBERG_MARQUARDT_POWER_FUNC = LEVENBERG_MARQUARDT_OPTIMIZATION2_FUNC(XDATA_MATRIX, YDATA_VECTOR, _
                PARAM_VECTOR, "CALL_POWER_OBJ_FUNC", _
                "CALL_POWER_GRAD_FUNC", 3, nLOOPS, tolerance, epsilon)

'--------------------------------------------------------------------------------
Case 4
'--------------------------------------------------------------------------------
    LEVENBERG_MARQUARDT_POWER_FUNC = LEVENBERG_MARQUARDT_OPTIMIZATION2_FUNC(XDATA_MATRIX, YDATA_VECTOR, _
                PARAM_VECTOR, "CALL_POWER_OBJ_FUNC", _
                "CALL_POWER_GRAD_FUNC", 4, nLOOPS, tolerance, epsilon)

'--------------------------------------------------------------------------------
Case Else
'--------------------------------------------------------------------------------
    LEVENBERG_MARQUARDT_POWER_FUNC = LEVENBERG_MARQUARDT_OPTIMIZATION2_FUNC(XDATA_MATRIX, YDATA_VECTOR, _
                PARAM_VECTOR, "CALL_POWER_OBJ_FUNC", _
                "CALL_POWER_GRAD_FUNC", 5, nLOOPS, tolerance, epsilon)

'--------------------------------------------------------------------------------
End Select
'--------------------------------------------------------------------------------


Exit Function
ERROR_LABEL:
LEVENBERG_MARQUARDT_POWER_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_POWER_OBJ_FUNC
'DESCRIPTION   :
'LIBRARY       : OPTIMIZATION
'GROUP         : NLR_FRAMES
'ID            : 016
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function CALL_POWER_OBJ_FUNC(ByRef XDATA_RNG As Variant, _
ByRef PARAM_RNG As Variant)

Dim i As Long
Dim NROWS As Long
Dim NSIZE As Long
Dim PARAM_VECTOR As Variant
Dim XDATA_MATRIX As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)
NSIZE = UBound(PARAM_VECTOR, 1)

XDATA_MATRIX = XDATA_RNG
If UBound(XDATA_MATRIX, 1) = 1 Then: XDATA_MATRIX = MATRIX_TRANSPOSE_FUNC(XDATA_MATRIX)

NROWS = UBound(XDATA_MATRIX, 1)

ReDim TEMP_MATRIX(1 To NROWS, 1 To 1)

For i = 1 To NROWS
    If XDATA_MATRIX(i, 1) > 0 Then
            TEMP_MATRIX(i, 1) = PARAM_VECTOR(1, 1) * XDATA_MATRIX(i, 1) ^ _
                                 PARAM_VECTOR(2, 1)
    End If
Next i

CALL_POWER_OBJ_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
CALL_POWER_OBJ_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_POWER_GRAD_FUNC
'DESCRIPTION   :
'LIBRARY       : OPTIMIZATION
'GROUP         : NLR_FRAMES
'ID            : 017
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function CALL_POWER_GRAD_FUNC(ByRef XDATA_RNG As Variant, _
ByRef PARAM_RNG As Variant)

Dim i As Long
Dim NROWS As Long
Dim NSIZE As Long
Dim PARAM_VECTOR As Variant
Dim XDATA_MATRIX As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)
NSIZE = UBound(PARAM_VECTOR, 1)

XDATA_MATRIX = XDATA_RNG
If UBound(XDATA_MATRIX, 1) = 1 Then: XDATA_MATRIX = MATRIX_TRANSPOSE_FUNC(XDATA_MATRIX)

NROWS = UBound(XDATA_MATRIX, 1)

ReDim TEMP_MATRIX(1 To NROWS, 1 To 2)

For i = 1 To NROWS
    If XDATA_MATRIX(i, 1) > 0 Then
        TEMP_MATRIX(i, 1) = XDATA_MATRIX(i, 1) ^ PARAM_VECTOR(2, 1)
        TEMP_MATRIX(i, 2) = (PARAM_VECTOR(1, 1) * XDATA_MATRIX(i, 1) ^ _
                            PARAM_VECTOR(2, 1)) * Log(XDATA_MATRIX(i, 1))
    End If
Next i

CALL_POWER_GRAD_FUNC = TEMP_MATRIX
Exit Function
ERROR_LABEL:
CALL_POWER_GRAD_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : LEVENBERG_MARQUARDT_LOGARITHM_FUNC
'DESCRIPTION   : Strictly related to the power fitting is the logarithmic regression.
'y = a * log(x) + b. Where "a" and "b" are parameters to determine.
'We performs this kind of regression when we have dataset sampled over
'a wide interval Range. For example the following dataset cames from the
'armonic analysis of a system. The vibration amplitude, was measured at
'10 different frequencies, from 0.1 KHz to about 2000 KHz.
'We usually plot this dataset with the help of a half-logarithm chart, and
'thus, it is reasonable to assume also a logarithm regression

'LIBRARY       : OPTIMIZATION
'GROUP         : NLR_FRAMES
'ID            : 018
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function LEVENBERG_MARQUARDT_LOGARITHM_FUNC(ByRef XDATA_RNG As Variant, _
ByRef YDATA_RNG As Variant, _
Optional ByRef PARAM_RNG As Variant, _
Optional ByVal nLOOPS As Long = 1000, _
Optional ByVal tolerance As Double = 10 ^ -10, _
Optional ByVal epsilon As Double = 10 ^ -6, _
Optional ByVal OUTPUT As Integer = 0)

'logarithm fitting

Dim i As Long
Dim NROWS As Long

Dim XDATA_MATRIX As Variant
Dim YDATA_VECTOR As Variant

Dim PARAM_VECTOR As Variant

Dim BTEMP_VECTOR As Variant

On Error GoTo ERROR_LABEL

XDATA_MATRIX = XDATA_RNG
If UBound(XDATA_MATRIX, 1) = 1 Then: XDATA_MATRIX = MATRIX_TRANSPOSE_FUNC(XDATA_MATRIX)

YDATA_VECTOR = YDATA_RNG
If UBound(YDATA_VECTOR, 1) = 1 Then: YDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(YDATA_VECTOR)

If UBound(XDATA_MATRIX, 1) <> UBound(YDATA_VECTOR, 1) Then: GoTo ERROR_LABEL
NROWS = UBound(XDATA_MATRIX, 1)

If IsArray(PARAM_RNG) = False Then 'parameter estimation
    
    ReDim ATEMP_VECTOR(1 To NROWS, 1 To 2)
    ReDim BTEMP_VECTOR(1 To NROWS, 1 To 1)
    
    For i = 1 To NROWS
        If XDATA_MATRIX(i, 1) <= 0 Then
           GoTo ERROR_LABEL
           'negative values in the dataset
        End If
        ATEMP_VECTOR(i, 1) = Log(XDATA_MATRIX(i, 1))
        ATEMP_VECTOR(i, 2) = 1
        BTEMP_VECTOR(i, 1) = YDATA_VECTOR(i, 1)
    Next i

    PARAM_VECTOR = MATRIX_LEAST_SQUARE_LINEAR_SYSTEM_FUNC(ATEMP_VECTOR, BTEMP_VECTOR, 0)

    
    PARAM_VECTOR = MATRIX_TRIANGULAR_LINEAR_SYSTEM_FUNC(PARAM_VECTOR)
    If PARAM_VECTOR(UBound(PARAM_VECTOR)) <> 0 Then
        PARAM_VECTOR = PARAM_VECTOR(LBound(PARAM_VECTOR))
    Else
        GoTo ERROR_LABEL
    End If
    
Else
    PARAM_VECTOR = PARAM_RNG
    If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)
    If UBound(PARAM_VECTOR, 1) <> 2 Then: GoTo ERROR_LABEL
End If

'PARAM_VECTOR --> [a,b]


'--------------------------------------------------------------------------------
Select Case OUTPUT
'--------------------------------------------------------------------------------
Case 0
'--------------------------------------------------------------------------------
    LEVENBERG_MARQUARDT_LOGARITHM_FUNC = LEVENBERG_MARQUARDT_OPTIMIZATION2_FUNC(XDATA_MATRIX, YDATA_VECTOR, _
                PARAM_VECTOR, "CALL_LOGARITHM_OBJ_FUNC", _
                "CALL_LOGARITHM_GRAD_FUNC", 0, nLOOPS, tolerance, epsilon)
'--------------------------------------------------------------------------------
Case 1
'--------------------------------------------------------------------------------
    LEVENBERG_MARQUARDT_LOGARITHM_FUNC = LEVENBERG_MARQUARDT_OPTIMIZATION2_FUNC(XDATA_MATRIX, YDATA_VECTOR, _
                PARAM_VECTOR, "CALL_LOGARITHM_OBJ_FUNC", _
                "CALL_LOGARITHM_GRAD_FUNC", 1, nLOOPS, tolerance, epsilon)

'--------------------------------------------------------------------------------
Case 2
'--------------------------------------------------------------------------------
    LEVENBERG_MARQUARDT_LOGARITHM_FUNC = LEVENBERG_MARQUARDT_OPTIMIZATION2_FUNC(XDATA_MATRIX, YDATA_VECTOR, _
                PARAM_VECTOR, "CALL_LOGARITHM_OBJ_FUNC", _
                "CALL_LOGARITHM_GRAD_FUNC", 2, nLOOPS, tolerance, epsilon)

'--------------------------------------------------------------------------------
Case 3
'--------------------------------------------------------------------------------
    LEVENBERG_MARQUARDT_LOGARITHM_FUNC = LEVENBERG_MARQUARDT_OPTIMIZATION2_FUNC(XDATA_MATRIX, YDATA_VECTOR, _
                PARAM_VECTOR, "CALL_LOGARITHM_OBJ_FUNC", _
                "CALL_LOGARITHM_GRAD_FUNC", 3, nLOOPS, tolerance, epsilon)

'--------------------------------------------------------------------------------
Case 4
'--------------------------------------------------------------------------------
    LEVENBERG_MARQUARDT_LOGARITHM_FUNC = LEVENBERG_MARQUARDT_OPTIMIZATION2_FUNC(XDATA_MATRIX, YDATA_VECTOR, _
                PARAM_VECTOR, "CALL_LOGARITHM_OBJ_FUNC", _
                "CALL_LOGARITHM_GRAD_FUNC", 4, nLOOPS, tolerance, epsilon)

'--------------------------------------------------------------------------------
Case Else
'--------------------------------------------------------------------------------
    LEVENBERG_MARQUARDT_LOGARITHM_FUNC = LEVENBERG_MARQUARDT_OPTIMIZATION2_FUNC(XDATA_MATRIX, YDATA_VECTOR, _
                PARAM_VECTOR, "CALL_LOGARITHM_OBJ_FUNC", _
                "CALL_LOGARITHM_GRAD_FUNC", 5, nLOOPS, tolerance, epsilon)

'--------------------------------------------------------------------------------
End Select
'--------------------------------------------------------------------------------


Exit Function
ERROR_LABEL:
LEVENBERG_MARQUARDT_LOGARITHM_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_LOGARITHM_OBJ_FUNC
'DESCRIPTION   :
'LIBRARY       : OPTIMIZATION
'GROUP         : NLR_FRAMES
'ID            : 019
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function CALL_LOGARITHM_OBJ_FUNC(ByRef XDATA_RNG As Variant, _
ByRef PARAM_RNG As Variant)

Dim i As Long
Dim NROWS As Long
Dim NSIZE As Long
Dim PARAM_VECTOR As Variant
Dim XDATA_MATRIX As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)
NSIZE = UBound(PARAM_VECTOR, 1)

XDATA_MATRIX = XDATA_RNG
If UBound(XDATA_MATRIX, 1) = 1 Then: XDATA_MATRIX = MATRIX_TRANSPOSE_FUNC(XDATA_MATRIX)

NROWS = UBound(XDATA_MATRIX, 1)

ReDim TEMP_MATRIX(1 To NROWS, 1 To 1)

For i = 1 To NROWS
    TEMP_MATRIX(i, 1) = PARAM_VECTOR(1, 1) * Log(XDATA_MATRIX(i, 1)) + _
                                PARAM_VECTOR(2, 1)
Next i

CALL_LOGARITHM_OBJ_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
CALL_LOGARITHM_OBJ_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_LOGARITHM_GRAD_FUNC
'DESCRIPTION   :
'LIBRARY       : OPTIMIZATION
'GROUP         : NLR_FRAMES
'ID            : 020
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function CALL_LOGARITHM_GRAD_FUNC(ByRef XDATA_RNG As Variant, _
ByRef PARAM_RNG As Variant)

Dim i As Long
Dim NROWS As Long
Dim NSIZE As Long
Dim PARAM_VECTOR As Variant
Dim XDATA_MATRIX As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)
NSIZE = UBound(PARAM_VECTOR, 1)

XDATA_MATRIX = XDATA_RNG
If UBound(XDATA_MATRIX, 1) = 1 Then: XDATA_MATRIX = MATRIX_TRANSPOSE_FUNC(XDATA_MATRIX)

NROWS = UBound(XDATA_MATRIX, 1)

ReDim TEMP_MATRIX(1 To NROWS, 1 To 2)

For i = 1 To NROWS
    TEMP_MATRIX(i, 1) = Log(XDATA_MATRIX(i, 1))
    TEMP_MATRIX(i, 2) = 1
Next i

CALL_LOGARITHM_GRAD_FUNC = TEMP_MATRIX
Exit Function
ERROR_LABEL:
CALL_LOGARITHM_GRAD_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : LEVENBERG_MARQUARDT_LOGISTIC_FUNC

'DESCRIPTION   : Logistic Cost fitting
'A method was needed for estimating the cash flows of engineering development
'projects undertaken by a certain company. One such project took 13 months to
'complete and the cumulative cost returns were collected throughout the life of
'the project. The accounts were closed 2 months after completion of the
'project, when the last bills were brought to account.

'The cumulative costs were collected at the end of each month up to the final
'fixed price of £1,000,000. The data therefore comprises 16 pairs of (end) of
'month numbers and the cumulative costs in £k.

'The cumulative logistic distribution function often fits data from growth
'situations that are limited by a finite resource. In this case, the costs are
'limited by the fixed price for the job. They grow slowly at first as just a
'few, then more and more people on the development team become involved. They
'then increase more rapidly as parts are bought in and manufacturing, assembly
'and test proceed, and then taper off as the manufacturing and development teams
'reduce with final evaluation and delivery to the customer, followed by
'settlement of the last bills from suppliers.

'The shape of the plot is typical of the cumulative expenditure for a fixed
'price project, and the sigmoidal form suggests that the logistic equation
'should fit the data. (Note the familiar slow start because engineers and
'draughtsmen were still involved with a different project!)

'References:
'http://www.cs.ubc.ca/spider/lowe/papers/pami91/pamilatex.html

'LIBRARY       : OPTIMIZATION
'GROUP         : NLR_FRAMES
'ID            : 021
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function LEVENBERG_MARQUARDT_LOGISTIC_FUNC(ByRef XDATA_RNG As Variant, _
ByRef YDATA_RNG As Variant, _
ByRef PARAM_RNG As Variant, _
Optional ByVal nLOOPS As Long = 1000, _
Optional ByVal tolerance As Double = 10 ^ -10, _
Optional ByVal epsilon As Double = 10 ^ -6, _
Optional ByVal OUTPUT As Integer = 0)

Dim NROWS As Long

Dim PARAM_VECTOR As Variant
Dim XDATA_MATRIX As Variant
Dim YDATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

XDATA_MATRIX = XDATA_RNG
If UBound(XDATA_MATRIX, 1) = 1 Then: XDATA_MATRIX = MATRIX_TRANSPOSE_FUNC(XDATA_MATRIX)

YDATA_VECTOR = YDATA_RNG
If UBound(YDATA_VECTOR, 1) = 1 Then: YDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(YDATA_VECTOR)

If UBound(XDATA_MATRIX, 1) <> UBound(YDATA_VECTOR, 1) Then: GoTo ERROR_LABEL
NROWS = UBound(XDATA_MATRIX, 1)

PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)
If UBound(PARAM_VECTOR, 1) <> 3 Then: GoTo ERROR_LABEL

'--------------------------------------------------------------------------------
Select Case OUTPUT
'--------------------------------------------------------------------------------
Case 0
'--------------------------------------------------------------------------------
    LEVENBERG_MARQUARDT_LOGISTIC_FUNC = LEVENBERG_MARQUARDT_OPTIMIZATION2_FUNC(XDATA_MATRIX, YDATA_VECTOR, _
                PARAM_VECTOR, "CALL_LOGISTIC_OBJ_FUNC", _
                "CALL_LOGISTIC_GRAD_FUNC", 0, nLOOPS, tolerance, epsilon)
'--------------------------------------------------------------------------------
Case 1
'--------------------------------------------------------------------------------
    LEVENBERG_MARQUARDT_LOGISTIC_FUNC = LEVENBERG_MARQUARDT_OPTIMIZATION2_FUNC(XDATA_MATRIX, YDATA_VECTOR, _
                PARAM_VECTOR, "CALL_LOGISTIC_OBJ_FUNC", _
                "CALL_LOGISTIC_GRAD_FUNC", 1, nLOOPS, tolerance, epsilon)

'--------------------------------------------------------------------------------
Case 2
'--------------------------------------------------------------------------------
    LEVENBERG_MARQUARDT_LOGISTIC_FUNC = LEVENBERG_MARQUARDT_OPTIMIZATION2_FUNC(XDATA_MATRIX, YDATA_VECTOR, _
                PARAM_VECTOR, "CALL_LOGISTIC_OBJ_FUNC", _
                "CALL_LOGISTIC_GRAD_FUNC", 2, nLOOPS, tolerance, epsilon)

'--------------------------------------------------------------------------------
Case 3
'--------------------------------------------------------------------------------
    LEVENBERG_MARQUARDT_LOGISTIC_FUNC = LEVENBERG_MARQUARDT_OPTIMIZATION2_FUNC(XDATA_MATRIX, YDATA_VECTOR, _
                PARAM_VECTOR, "CALL_LOGISTIC_OBJ_FUNC", _
                "CALL_LOGISTIC_GRAD_FUNC", 3, nLOOPS, tolerance, epsilon)

'--------------------------------------------------------------------------------
Case 4
'--------------------------------------------------------------------------------
    LEVENBERG_MARQUARDT_LOGISTIC_FUNC = LEVENBERG_MARQUARDT_OPTIMIZATION2_FUNC(XDATA_MATRIX, YDATA_VECTOR, _
                PARAM_VECTOR, "CALL_LOGISTIC_OBJ_FUNC", _
                "CALL_LOGISTIC_GRAD_FUNC", 4, nLOOPS, tolerance, epsilon)

'--------------------------------------------------------------------------------
Case Else
'--------------------------------------------------------------------------------
    LEVENBERG_MARQUARDT_LOGISTIC_FUNC = LEVENBERG_MARQUARDT_OPTIMIZATION2_FUNC(XDATA_MATRIX, YDATA_VECTOR, _
                PARAM_VECTOR, "CALL_LOGISTIC_OBJ_FUNC", _
                "CALL_LOGISTIC_GRAD_FUNC", 5, nLOOPS, tolerance, epsilon)

'--------------------------------------------------------------------------------
End Select
'--------------------------------------------------------------------------------


Exit Function
ERROR_LABEL:
LEVENBERG_MARQUARDT_LOGISTIC_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_LOGISTIC_OBJ_FUNC
'DESCRIPTION   :
'LIBRARY       : OPTIMIZATION
'GROUP         : NLR_FRAMES
'ID            : 022
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function CALL_LOGISTIC_OBJ_FUNC(ByRef XDATA_RNG As Variant, _
ByRef PARAM_RNG As Variant)

Dim i As Long
Dim NROWS As Long
Dim NSIZE As Long
Dim PARAM_VECTOR As Variant
Dim XDATA_MATRIX As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)
NSIZE = UBound(PARAM_VECTOR, 1)

XDATA_MATRIX = XDATA_RNG
If UBound(XDATA_MATRIX, 1) = 1 Then: XDATA_MATRIX = MATRIX_TRANSPOSE_FUNC(XDATA_MATRIX)

NROWS = UBound(XDATA_MATRIX, 1)

ReDim TEMP_MATRIX(1 To NROWS, 1 To 1)

For i = 1 To NROWS
'------------------------------------------------------------------------------------
'    TEMP_MATRIX(i, 1) = PARAM_VECTOR(1, 1) * PARAM_VECTOR(2, 1) / _
                        (PARAM_VECTOR(1, 1) + (PARAM_VECTOR(2, 1) - _
                                PARAM_VECTOR(1, 1)) * Exp(-PARAM_VECTOR(3, 1) * _
                                        XDATA_MATRIX(i, 1)))
'------------------------------------------------------------------------------------
    
    TEMP_MATRIX(i, 1) = PARAM_VECTOR(1, 1) / (1 + Exp(PARAM_VECTOR(2, 1) - _
                         PARAM_VECTOR(3, 1) * XDATA_MATRIX(i, 1)))
Next i

CALL_LOGISTIC_OBJ_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
CALL_LOGISTIC_OBJ_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_LOGISTIC_GRAD_FUNC
'DESCRIPTION   :
'LIBRARY       : OPTIMIZATION
'GROUP         : NLR_FRAMES
'ID            : 023
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function CALL_LOGISTIC_GRAD_FUNC(ByRef XDATA_RNG As Variant, _
ByRef PARAM_RNG As Variant)

Dim i As Long
Dim NROWS As Long
Dim NSIZE As Long

'Dim ATEMP_VAL As Double
'Dim BTEMP_VAL As Double

Dim PARAM_VECTOR As Variant
Dim XDATA_MATRIX As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)
NSIZE = UBound(PARAM_VECTOR, 1)

XDATA_MATRIX = XDATA_RNG
If UBound(XDATA_MATRIX, 1) = 1 Then: XDATA_MATRIX = MATRIX_TRANSPOSE_FUNC(XDATA_MATRIX)

NROWS = UBound(XDATA_MATRIX, 1)

ReDim TEMP_MATRIX(1 To NROWS, 1 To 3)

For i = 1 To NROWS
'------------------------------------------------------------------------------------
'    ATEMP_VAL = Exp(-PARAM_VECTOR(3, 1) * XDATA_MATRIX(i, 1))
 '   BTEMP_VAL = (PARAM_VECTOR(1, 1) + (PARAM_VECTOR(2, 1) - _
                        PARAM_VECTOR(1, 1)) * ATEMP_VAL)
 '   TEMP_MATRIX(i, 1) = (BTEMP_VAL - (1 - ATEMP_VAL) * _
                            PARAM_VECTOR(1, 1)) * _
                            PARAM_VECTOR(2, 1) / BTEMP_VAL ^ 2
  '  TEMP_MATRIX(i, 2) = (BTEMP_VAL - ATEMP_VAL * PARAM_VECTOR(2, 1)) * _
                            PARAM_VECTOR(1, 1) / BTEMP_VAL ^ 2
   ' TEMP_MATRIX(i, 3) = XDATA_MATRIX(i, 1) * (PARAM_VECTOR(2, 1) - _
                            PARAM_VECTOR(1, 1)) * ATEMP_VAL * PARAM_VECTOR(1, 1) * _
                                    PARAM_VECTOR(2, 1) / BTEMP_VAL ^ 2
'------------------------------------------------------------------------------------

    TEMP_MATRIX(i, 1) = 1 / (1 + Exp(PARAM_VECTOR(2, 1) - _
                        PARAM_VECTOR(3, 1) * XDATA_MATRIX(i, 1)))
    
    TEMP_MATRIX(i, 2) = -PARAM_VECTOR(1, 1) / (1 + Exp(PARAM_VECTOR(2, 1) - _
                        PARAM_VECTOR(3, 1) * XDATA_MATRIX(i, 1))) ^ 2 * _
                        Exp(PARAM_VECTOR(2, 1) - PARAM_VECTOR(3, 1) * _
                        XDATA_MATRIX(i, 1))
    
    TEMP_MATRIX(i, 3) = PARAM_VECTOR(1, 1) / (1 + Exp(PARAM_VECTOR(2, 1) - _
                         PARAM_VECTOR(3, 1) * XDATA_MATRIX(i, 1))) ^ 2 * _
                         XDATA_MATRIX(i, 1) * Exp(PARAM_VECTOR(2, 1) - _
                         PARAM_VECTOR(3, 1) * XDATA_MATRIX(i, 1))

Next i

CALL_LOGISTIC_GRAD_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
CALL_LOGISTIC_GRAD_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : LEVENBERG_MARQUARDT_BI_DIMEN_GAUSS_FUNC

'DESCRIPTION   : Bi-Dimensional Gauss fitting
'The Gaussian function accounts for distribution anomalies,
'for example the bombardment of molecules to elicit ions in mass
'spectometry. The kinetic energy of the molecules before ionization
'would ideally be zero in reference to the direction of acceleration
'through the magnetic sector. Though if the respective molecule's kinetic
'energy was negative prior to the moment of ionization by the transferred
'kinetic energy of the electron, the resultant velocity at the detector
'would be slower than predicted. Vice versa for a molecule maintaining a
'relative positive velocity prior to the moment of ionization experiences a
'faster velocity than predicted for it's weight. The effects are represented
'in a bell curve whereby the bulk lie within the predicted domain, and
'deviations from which diminish exponentially in either direction from that
'point. The aforementioned accounts for a single prescribed distribution,
'though the effects may be complexed by multiple overlapping bell curves.
'For visual, imagine a distributed anomaly falling within a neighboring bell,
'thereby distorting the seen data.

'The integral of the Gaussian function is the error function.
'Gaussian functions appear in many contexts in the natural sciences, the
'social sciences, mathematics, and engineering. Some examples include:
'In statistics and probability theory, Gaussian functions appear as the
'density function of the normal distribution, which is a limiting probability
'distribution of complicated sums, according to the central limit theorem.
'A Gaussian function is the wave function of the ground state of the quantum
'harmonic oscillator. The molecular orbitals used in computational chemistry
'can be linear combinations of Gaussian functions called Gaussian orbitals
'Mathematically, the Gaussian function plays an important role in the definition
'of the Hermite polynomials. Consequently, Gaussian functions are also associated
'with the vacuum state in quantum field theory. Gaussian beams are used in optical
'and microwave systems, Gaussian functions are used as smoothing kernels for
'generating multi-scale representations in computer vision and image processing
'Specifically, derivatives of Gaussians are used as a basis for defining a large
'number of types of visual operations. Gaussian functions are used in some
'types of artificial neural networks.

'LIBRARY       : OPTIMIZATION
'GROUP         : NLR_FRAMES
'ID            : 024
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function LEVENBERG_MARQUARDT_BI_DIMEN_GAUSS_FUNC(ByRef XDATA_RNG As Variant, _
ByRef YDATA_RNG As Variant, _
ByRef PARAM_RNG As Variant, _
Optional ByVal nLOOPS As Long = 1000, _
Optional ByVal tolerance As Double = 10 ^ -10, _
Optional ByVal epsilon As Double = 10 ^ -6, _
Optional ByVal OUTPUT As Integer = 0)

Dim NROWS As Long

Dim PARAM_VECTOR As Variant
Dim XDATA_MATRIX As Variant
Dim YDATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

XDATA_MATRIX = XDATA_RNG
If UBound(XDATA_MATRIX, 1) = 1 Then: XDATA_MATRIX = MATRIX_TRANSPOSE_FUNC(XDATA_MATRIX)

YDATA_VECTOR = YDATA_RNG
If UBound(YDATA_VECTOR, 1) = 1 Then: YDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(YDATA_VECTOR)

If UBound(XDATA_MATRIX, 1) <> UBound(YDATA_VECTOR, 1) Then: GoTo ERROR_LABEL
NROWS = UBound(XDATA_MATRIX, 1)

PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)
If UBound(PARAM_VECTOR, 1) <> 4 Then: GoTo ERROR_LABEL

'--------------------------------------------------------------------------------
Select Case OUTPUT
'--------------------------------------------------------------------------------
Case 0
'--------------------------------------------------------------------------------
    LEVENBERG_MARQUARDT_BI_DIMEN_GAUSS_FUNC = LEVENBERG_MARQUARDT_OPTIMIZATION2_FUNC(XDATA_MATRIX, YDATA_VECTOR, _
                PARAM_VECTOR, "CALL_BI_DIMEN_GAUSS_OBJ_FUNC", _
                "CALL_BI_DIMEN_GAUSS_GRAD_FUNC", 0, nLOOPS, tolerance, epsilon)
'--------------------------------------------------------------------------------
Case 1
'--------------------------------------------------------------------------------
    LEVENBERG_MARQUARDT_BI_DIMEN_GAUSS_FUNC = LEVENBERG_MARQUARDT_OPTIMIZATION2_FUNC(XDATA_MATRIX, YDATA_VECTOR, _
                PARAM_VECTOR, "CALL_BI_DIMEN_GAUSS_OBJ_FUNC", _
                "CALL_BI_DIMEN_GAUSS_GRAD_FUNC", 1, nLOOPS, tolerance, epsilon)

'--------------------------------------------------------------------------------
Case 2
'--------------------------------------------------------------------------------
    LEVENBERG_MARQUARDT_BI_DIMEN_GAUSS_FUNC = LEVENBERG_MARQUARDT_OPTIMIZATION2_FUNC(XDATA_MATRIX, YDATA_VECTOR, _
                PARAM_VECTOR, "CALL_BI_DIMEN_GAUSS_OBJ_FUNC", _
                "CALL_BI_DIMEN_GAUSS_GRAD_FUNC", 2, nLOOPS, tolerance, epsilon)

'--------------------------------------------------------------------------------
Case 3
'--------------------------------------------------------------------------------
    LEVENBERG_MARQUARDT_BI_DIMEN_GAUSS_FUNC = LEVENBERG_MARQUARDT_OPTIMIZATION2_FUNC(XDATA_MATRIX, YDATA_VECTOR, _
                PARAM_VECTOR, "CALL_BI_DIMEN_GAUSS_OBJ_FUNC", _
                "CALL_BI_DIMEN_GAUSS_GRAD_FUNC", 3, nLOOPS, tolerance, epsilon)

'--------------------------------------------------------------------------------
Case 4
'--------------------------------------------------------------------------------
    LEVENBERG_MARQUARDT_BI_DIMEN_GAUSS_FUNC = LEVENBERG_MARQUARDT_OPTIMIZATION2_FUNC(XDATA_MATRIX, YDATA_VECTOR, _
                PARAM_VECTOR, "CALL_BI_DIMEN_GAUSS_OBJ_FUNC", _
                "CALL_BI_DIMEN_GAUSS_GRAD_FUNC", 4, nLOOPS, tolerance, epsilon)

'--------------------------------------------------------------------------------
Case Else
'--------------------------------------------------------------------------------
    LEVENBERG_MARQUARDT_BI_DIMEN_GAUSS_FUNC = LEVENBERG_MARQUARDT_OPTIMIZATION2_FUNC(XDATA_MATRIX, YDATA_VECTOR, _
                PARAM_VECTOR, "CALL_BI_DIMEN_GAUSS_OBJ_FUNC", _
                "CALL_BI_DIMEN_GAUSS_GRAD_FUNC", 5, nLOOPS, tolerance, epsilon)

'--------------------------------------------------------------------------------
End Select
'--------------------------------------------------------------------------------


Exit Function
ERROR_LABEL:
LEVENBERG_MARQUARDT_BI_DIMEN_GAUSS_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_BI_DIMEN_GAUSS_OBJ_FUNC
'DESCRIPTION   :
'LIBRARY       : OPTIMIZATION
'GROUP         : NLR_FRAMES
'ID            : 025
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function CALL_BI_DIMEN_GAUSS_OBJ_FUNC(ByRef XDATA_RNG As Variant, _
ByRef PARAM_RNG As Variant)

Dim i As Long
Dim NROWS As Long
Dim NSIZE As Long
Dim PARAM_VECTOR As Variant
Dim XDATA_MATRIX As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)
NSIZE = UBound(PARAM_VECTOR, 1)

XDATA_MATRIX = XDATA_RNG
If UBound(XDATA_MATRIX, 1) = 1 Then: XDATA_MATRIX = MATRIX_TRANSPOSE_FUNC(XDATA_MATRIX)

NROWS = UBound(XDATA_MATRIX, 1)

ReDim TEMP_MATRIX(1 To NROWS, 1 To 1)

For i = 1 To NROWS
    TEMP_MATRIX(i, 1) = PARAM_VECTOR(4, 1) * Exp(-(PARAM_VECTOR(1, 1) * _
                    XDATA_MATRIX(i, 1) ^ 2 + PARAM_VECTOR(2, 1) * _
                    XDATA_MATRIX(i, 2) ^ 2 + PARAM_VECTOR(3, 1) * _
                    XDATA_MATRIX(i, 1) * XDATA_MATRIX(i, 2)))
Next i

CALL_BI_DIMEN_GAUSS_OBJ_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
CALL_BI_DIMEN_GAUSS_OBJ_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_BI_DIMEN_GAUSS_GRAD_FUNC
'DESCRIPTION   :
'LIBRARY       : OPTIMIZATION
'GROUP         : NLR_FRAMES
'ID            : 026
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function CALL_BI_DIMEN_GAUSS_GRAD_FUNC(ByRef XDATA_RNG As Variant, _
ByRef PARAM_RNG As Variant)

Dim i As Long
Dim NROWS As Long
Dim NSIZE As Long

Dim TEMP_VALUE As Double

Dim PARAM_VECTOR As Variant
Dim XDATA_MATRIX As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then: _
    PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)
NSIZE = UBound(PARAM_VECTOR, 1)

XDATA_MATRIX = XDATA_RNG
If UBound(XDATA_MATRIX, 1) = 1 Then: _
    XDATA_MATRIX = MATRIX_TRANSPOSE_FUNC(XDATA_MATRIX)

NROWS = UBound(XDATA_MATRIX, 1)

ReDim TEMP_MATRIX(1 To NROWS, 1 To 4)

For i = 1 To NROWS
                
    TEMP_VALUE = Exp(-(PARAM_VECTOR(1, 1) * XDATA_MATRIX(i, 1) ^ 2 _
                    + PARAM_VECTOR(2, 1) * XDATA_MATRIX(i, 2) ^ 2 + _
                    PARAM_VECTOR(3, 1) * XDATA_MATRIX(i, 1) * XDATA_MATRIX(i, 2)))
                
    TEMP_MATRIX(i, 1) = -PARAM_VECTOR(4, 1) * _
                                XDATA_MATRIX(i, 1) ^ 2 * TEMP_VALUE
                
    TEMP_MATRIX(i, 2) = -PARAM_VECTOR(4, 1) * _
                                XDATA_MATRIX(i, 2) ^ 2 * TEMP_VALUE
                
    TEMP_MATRIX(i, 3) = -PARAM_VECTOR(4, 1) * XDATA_MATRIX(i, 1) * _
                                XDATA_MATRIX(i, 2) * TEMP_VALUE
                
    TEMP_MATRIX(i, 4) = TEMP_VALUE
Next i

CALL_BI_DIMEN_GAUSS_GRAD_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
CALL_BI_DIMEN_GAUSS_GRAD_FUNC = Err.number
End Function
