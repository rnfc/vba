Attribute VB_Name = "POLYNOMIAL_MOMENTS_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_EXTRACT_MOMENTS_FUNC

'DESCRIPTION   : This function extracts the moments from a given dataset xy
'using the Cavalieri-Simpson quadrature formula
'Integral( f(x)*x^m, a<x<b ) = M(m)  with m=0,1,2...22

'Method: Finding f(x) as expansion of Legendre polynomials that can be
'successfully used for fitting smooth regular functions using up to 10 to
'20 moments.

'http://digilander.libero.it/foxes/poly/Moments_Regression.pdf

'LIBRARY       : POLYNOMIAL
'GROUP         : MOMENTS
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 08/29/2010
'************************************************************************************
'************************************************************************************

Function MATRIX_EXTRACT_MOMENTS_FUNC(ByRef XDATA_RNG As Variant, _
ByRef YDATA_RNG As Variant, _
Optional ByVal NDEG As Long = 3)
'Moments Spectrum:
'The moment regression means the objective of finding a function f(x) from
'the knowledge of its moments. In other words, givens a set of moments we
'want to search a function f(x), if exists, that best fits the given moments
'in the specified integration interval. The numerical solution of this problem
'is not simple. Using theoretic formulas can easily fail for high complexity
'or instability (see the moments theorem).

Dim i As Long
Dim j As Long
Dim NROWS As Long

Dim MIN_VAL As Double
Dim MAX_VAL As Double
Dim D_VAL As Double

Dim TEMP1_SUM As Double
Dim TEMP2_SUM As Double

Dim XDATA_VECTOR As Variant
Dim YDATA_VECTOR As Variant

Dim TEMP_VECTOR As Variant
Dim MOMENTS_VECTOR As Variant

Dim epsilon As Double

On Error GoTo ERROR_LABEL

epsilon = 2 * 10 ^ -15

XDATA_VECTOR = XDATA_RNG
If UBound(XDATA_VECTOR, 1) = 1 Then
    XDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(XDATA_VECTOR)
End If

YDATA_VECTOR = YDATA_RNG
If UBound(YDATA_VECTOR, 1) = 1 Then
    YDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(YDATA_VECTOR)
End If
If UBound(XDATA_VECTOR, 1) <> UBound(YDATA_VECTOR, 1) Then: GoTo ERROR_LABEL

NROWS = UBound(XDATA_VECTOR, 1)
If NDEG > 20 Then: GoTo ERROR_LABEL
If NROWS Mod 2 = 0 Then: GoTo ERROR_LABEL 'points must be odd

MIN_VAL = 2 ^ 52: MAX_VAL = -2 ^ 52
For i = 1 To NROWS
    If XDATA_VECTOR(i, 1) < MIN_VAL Then: MIN_VAL = XDATA_VECTOR(i, 1)
    If XDATA_VECTOR(i, 1) > MAX_VAL Then: MAX_VAL = XDATA_VECTOR(i, 1)
Next i

j = 0
D_VAL = (MAX_VAL - MIN_VAL) / (NROWS - 1)

ReDim MOMENTS_VECTOR(0 To NDEG - 1, 1 To 1)
TEMP_VECTOR = YDATA_VECTOR

Do
    'compute the j-th moments
    TEMP1_SUM = 0: TEMP2_SUM = 0
    For i = 2 To NROWS - 1 Step 2
        TEMP1_SUM = TEMP1_SUM + TEMP_VECTOR(i, 1)
    Next i
    For i = 3 To NROWS - 1 Step 2
        TEMP2_SUM = TEMP2_SUM + TEMP_VECTOR(i, 1)
    Next i
    MOMENTS_VECTOR(j, 1) = D_VAL / 3 * (TEMP_VECTOR(1, 1) + 4 * TEMP1_SUM + 2 * TEMP2_SUM + TEMP_VECTOR(NROWS, 1))
    
    If Abs(MOMENTS_VECTOR(j, 1)) < epsilon Then: MOMENTS_VECTOR(j, 1) = 0
    j = j + 1
    If j = NDEG Then Exit Do
    For i = 1 To NROWS
        TEMP_VECTOR(i, 1) = TEMP_VECTOR(i, 1) * XDATA_VECTOR(i, 1)
    Next i
Loop

MATRIX_EXTRACT_MOMENTS_FUNC = MOMENTS_VECTOR

Exit Function
ERROR_LABEL:
MATRIX_EXTRACT_MOMENTS_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : POLYNOMIAL_ORTHOGONAL_MOMENTS_REGRESSION_FUNC

'DESCRIPTION   : This routine performs the moments regression using the
'orthogonal polynomials polynomial order < 23. Max moments = 23
'Integral( f(x)*x^m, a<x<b ) = M(m)  with m=0,1,2...22
'where f(x)= c0+c1*x+c2*x^2+...cm*x^m

'LIBRARY       : POLYNOMIAL
'GROUP         : MOMENTS
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 08/29/2010
'************************************************************************************
'************************************************************************************

Function POLYNOMIAL_ORTHOGONAL_MOMENTS_REGRESSION_FUNC( _
ByRef MOMENTS_RNG As Variant, _
ByVal MIN_VAL As Double, _
ByVal MAX_VAL As Double, _
Optional ByVal POLYN_TYPE As Integer = 0, _
Optional ByVal LOWER_VAL As Double = 0, _
Optional ByVal UPPER_VAL As Double = 1)

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long

Dim NDEG As Long
Dim NROWS As Long

Dim R_VAL As Double
Dim C_VAL As Double
Dim MOMENTS_VECTOR As Variant
Dim COEFFICIENTS_VECTOR As Variant

Dim XDATA_ARR  As Variant
Dim MOMENTS_ARR As Variant

Dim TEMP1_MATRIX  As Variant
Dim TEMP2_MATRIX As Variant

Dim epsilon As Double

On Error GoTo ERROR_LABEL
'-----------------------------------------------------------------------------------------------------
epsilon = 2 * 10 ^ -15
'-----------------------------------------------------------------------------------------------------
MOMENTS_VECTOR = MOMENTS_RNG
If UBound(MOMENTS_VECTOR, 1) = 1 Then
    MOMENTS_VECTOR = MATRIX_TRANSPOSE_FUNC(MOMENTS_VECTOR)
End If
NROWS = UBound(MOMENTS_VECTOR)
NDEG = NROWS - 1  'max legendre poly NDEG

ReDim TEMP1_MATRIX(1 To NROWS, 1 To NROWS)
ReDim TEMP2_MATRIX(0 To NDEG, 0 To NDEG)

'-----------------------------------------------------------------------------------------------------
'build the coefficient matrix
'-----------------------------------------------------------------------------------------------------
TEMP2_MATRIX(0, 0) = 1
For h = 1 To NDEG 'Poly Builder
    COEFFICIENTS_VECTOR = POLYNOMIAL_TYPE_COEFFICIENTS_FUNC(h, POLYN_TYPE, LOWER_VAL, UPPER_VAL)
    For j = 0 To h
        'takes the i-th NDEG of polynomial
        TEMP2_MATRIX(h, j) = COEFFICIENTS_VECTOR(j, 1)
    Next j
Next h
'-----------------------------------------------------------------------------------------------------
'build the system solving matrix
'-----------------------------------------------------------------------------------------------------
For i = 1 To NROWS
    For j = 1 To i
        If (i + j) Mod 2 = 0 Then
            For k = 0 To j - 1
                If (k + i) Mod 2 <> 0 Then
                    TEMP1_MATRIX(i, j) = TEMP1_MATRIX(i, j) + TEMP2_MATRIX(j - 1, k) * 2 / (k + i)
                End If
            Next k
        End If
    Next j
Next i
'-----------------------------------------------------------------------------------------------------
'moments normalization
'-----------------------------------------------------------------------------------------------------
ReDim MOMENTS_ARR(0 To NDEG)
If MIN_VAL <> -1 Or MAX_VAL <> 1 Then
'-----------------------------------------------------------------------------------------------------
' This routine converts the moments from a given interval [a, b] to [-1, 1]
' MOMENTS_VECTOR =  raw moments in the interval [a, b]
' MOMENTS_ARR = normalized moments in the interval [-1, 1] (normalized)
    R_VAL = (MIN_VAL + MAX_VAL) / (MAX_VAL - MIN_VAL)
    For i = 0 To NDEG
        MOMENTS_ARR(i) = MOMENTS_VECTOR(i + 1, 1) * (2 / (MAX_VAL - MIN_VAL)) ^ (i + 1)
        For j = 1 To i
            C_VAL = COMBINATIONS_FUNC(i, j) 'WorksheetFunction.Combin(i, j)
            MOMENTS_ARR(i) = MOMENTS_ARR(i) - C_VAL * R_VAL ^ j * MOMENTS_ARR(i - j)
        Next j
    Next i
    For i = 0 To NDEG
        If Abs(MOMENTS_ARR(i)) < epsilon * (1 + Abs(MOMENTS_VECTOR(i + 1, 1))) Then MOMENTS_ARR(i) = 0
    Next i
'-----------------------------------------------------------------------------------------------------
Else
'-----------------------------------------------------------------------------------------------------
    For i = 0 To NDEG: MOMENTS_ARR(i) = MOMENTS_VECTOR(i + 1, 1): Next i
'-----------------------------------------------------------------------------------------------------
End If
'-----------------------------------------------------------------------------------------------------
'solve the triangular system with forward substitution
ReDim XDATA_ARR(1 To NROWS)
For i = 1 To NROWS
    For k = 1 To i - 1
        XDATA_ARR(i) = XDATA_ARR(i) - TEMP1_MATRIX(i, k) * XDATA_ARR(k)
    Next k
    XDATA_ARR(i) = (XDATA_ARR(i) + MOMENTS_ARR(i - 1)) / TEMP1_MATRIX(i, i)
Next i
'-----------------------------------------------------------------------------------------------------
'build the final regression polynomial
ReDim COEFFICIENTS_VECTOR(1 To NDEG + 1, 1 To 1)
For i = 0 To NDEG
    For k = 0 To NDEG
        COEFFICIENTS_VECTOR(i + 1, 1) = COEFFICIENTS_VECTOR(i + 1, 1) + XDATA_ARR(k + 1) * TEMP2_MATRIX(k, i)
    Next k
Next i
'-----------------------------------------------------------------------------------------------------
POLYNOMIAL_ORTHOGONAL_MOMENTS_REGRESSION_FUNC = COEFFICIENTS_VECTOR

Exit Function
ERROR_LABEL:
POLYNOMIAL_ORTHOGONAL_MOMENTS_REGRESSION_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : POLYNOMIAL_LEGENDRE_MOMENT_REGRESSION_SAMPLING_FUNC

'DESCRIPTION   : This routine tabulates the moments regression using the
'Legendre polynomials yi = p(xi) , xi=xmin+(b-a)/(pmax-1)*i  i=0,1, pmax-1
'input a, b = sampling interval [a, b]
'input pmax = max number of samples
'input coef = regression polynomial coefficients
'output Dxy = array of data sampled [xi, yi]

'LIBRARY       : POLYNOMIAL
'GROUP         : MOMENTS
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 08/29/2010
'************************************************************************************
'************************************************************************************

Function POLYNOMIAL_LEGENDRE_MOMENT_REGRESSION_SAMPLING_FUNC( _
ByRef COEF_RNG As Variant, _
ByVal MIN_VAL As Double, _
ByVal MAX_VAL As Double, _
Optional ByVal NBINS As Long = 100)

Dim i As Long
Dim j As Long

Dim Y_VAL As Double

Dim TEMP1_VAL As Double
Dim TEMP2_VAL As Double
Dim TEMP3_VAL As Double
Dim TEMP4_VAL As Double

Dim COEF_VECTOR As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

COEF_VECTOR = COEF_RNG
If UBound(COEF_VECTOR, 1) = 1 Then
    COEF_VECTOR = MATRIX_TRANSPOSE_FUNC(COEF_VECTOR)
End If

ReDim TEMP_MATRIX(0 To NBINS, 1 To 2)
TEMP_MATRIX(0, 1) = "X VAR"
TEMP_MATRIX(0, 2) = "Y VAR"

TEMP1_VAL = 2 / (NBINS - 1)
TEMP2_VAL = (MAX_VAL + MIN_VAL) / 2
TEMP3_VAL = (MAX_VAL - MIN_VAL) / 2

For i = 1 To NBINS
    TEMP4_VAL = (i - 1) * TEMP1_VAL - 1
    TEMP_MATRIX(i, 1) = TEMP3_VAL * TEMP4_VAL + TEMP2_VAL
    Y_VAL = 0     'Horner-Ruffini algorithm for polynomial computing
    For j = UBound(COEF_VECTOR) To LBound(COEF_VECTOR) Step -1
        Y_VAL = Y_VAL * TEMP4_VAL + COEF_VECTOR(j, 1)
    Next j
    TEMP_MATRIX(i, 2) = Y_VAL
Next i
POLYNOMIAL_LEGENDRE_MOMENT_REGRESSION_SAMPLING_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
POLYNOMIAL_LEGENDRE_MOMENT_REGRESSION_SAMPLING_FUNC = Err.number
End Function
