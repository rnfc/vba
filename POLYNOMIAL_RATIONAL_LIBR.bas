Attribute VB_Name = "POLYNOMIAL_RATIONAL_LIBR"

'-------------------------------------------------------------------------------------
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'-------------------------------------------------------------------------------------

'************************************************************************************
'************************************************************************************
'FUNCTION      : RATIONAL_POLYNOMIAL_QUOTIENT_FUNC

'DESCRIPTION   : Compute the quotient of 2 polynomials a(x)/b(x)
' q(x)= (a0+a1*x+a2*x^2+...+an*x^n)/(b0+b1*x+b2*x^2+...+x^m)
'na = degree of a(x)
'nb = degree of b(x)
'c = [a0, a1...an, b0, b1...bm-1]
'remember that the denominator polynomial is monic, i.e. bm =1

'LIBRARY       : POLYNOMIAL
'GROUP         : RATIONAL
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function RATIONAL_POLYNOMIAL_QUOTIENT_FUNC(ByRef XDATA_RNG As Variant, _
ByRef COEF_RNG As Variant, _
ByVal DEG_NUMER As Long, _
ByVal DEG_DENOM As Long)

Dim i As Long

Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double

Dim ATEMP_VECTOR As Variant
Dim BTEMP_VECTOR As Variant

Dim COEF_VECTOR As Variant

Dim DATA_VECTOR As Variant
Dim YTEMP_VECTOR As Variant

On Error GoTo ERROR_LABEL

COEF_VECTOR = COEF_RNG
DATA_VECTOR = XDATA_RNG

'split coefficients and constant
ReDim ATEMP_VECTOR(0 To DEG_NUMER, 1 To 1)
ReDim BTEMP_VECTOR(0 To DEG_DENOM, 1 To 1)

For i = 0 To DEG_NUMER
    ATEMP_VECTOR(i, 1) = COEF_VECTOR(i + 1, 1)
Next i
For i = 0 To DEG_DENOM - 1
    BTEMP_VECTOR(i, 1) = COEF_VECTOR(i + DEG_NUMER + 2, 1)
Next i

BTEMP_VECTOR(DEG_DENOM, 1) = 1 'evaluation begins
ReDim YTEMP_VECTOR(1 To UBound(DATA_VECTOR, 1), 1 To 1)
For i = 1 To UBound(DATA_VECTOR, 1)
    ATEMP_VAL = RATIONAL_POLYNOMIAL_EVALUATE_FUNC(DATA_VECTOR(i, 1), ATEMP_VECTOR, 0)
    BTEMP_VAL = RATIONAL_POLYNOMIAL_EVALUATE_FUNC(DATA_VECTOR(i, 1), BTEMP_VECTOR, 0)
    If BTEMP_VAL <> 0 Then
        YTEMP_VECTOR(i, 1) = ATEMP_VAL / BTEMP_VAL
    End If
Next i

RATIONAL_POLYNOMIAL_QUOTIENT_FUNC = YTEMP_VECTOR

Exit Function
ERROR_LABEL:
RATIONAL_POLYNOMIAL_QUOTIENT_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : RATIONAL_POLYNOMIAL_GRADIENT_FUNC

'DESCRIPTION   : Compute the derivatives of 2 polynomials quotient a(x)/b(x)
' q(x)= (a0+a1*x+a2*x^2+...+an*x^n)/(b0+b1*x+b2*x^2+...+x^m)
'respect to coefficients K, a0, a1, etc.
'na = degree of a(x)
'nb = degree of b(x)
'c = [a0, a1...an, b0, b1...bm-1]
'remember that the denominator polynomial is monic, i.e. bm =1

'LIBRARY       : POLYNOMIAL
'GROUP         : RATIONAL
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function RATIONAL_POLYNOMIAL_GRADIENT_FUNC(ByRef XDATA_RNG As Variant, _
ByRef COEF_RNG As Variant, _
ByVal DEG_NUMER As Long, _
ByVal DEG_DENOM As Long)

Dim i As Long
Dim j As Long

Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double

Dim GRAD_VECTOR As Variant
Dim COEF_VECTOR As Variant

Dim DATA_VECTOR As Variant

Dim ATEMP_VECTOR As Variant
Dim BTEMP_VECTOR As Variant

On Error GoTo ERROR_LABEL

COEF_VECTOR = COEF_RNG
DATA_VECTOR = XDATA_RNG

'split coefficients and constant
ReDim ATEMP_VECTOR(0 To DEG_NUMER, 1 To 1)
ReDim BTEMP_VECTOR(0 To DEG_DENOM, 1 To 1)

For i = 0 To DEG_NUMER
    ATEMP_VECTOR(i, 1) = COEF_VECTOR(i + 1, 1)
Next i

For i = 0 To DEG_DENOM - 1
    BTEMP_VECTOR(i, 1) = COEF_VECTOR(i + DEG_NUMER + 2, 1)
Next i

BTEMP_VECTOR(DEG_DENOM, 1) = 1 'evaluation begins
ReDim GRAD_VECTOR(1 To UBound(DATA_VECTOR, 1), _
                  1 To (DEG_DENOM + DEG_NUMER + 1))

For i = 1 To UBound(DATA_VECTOR, 1)
    ATEMP_VAL = RATIONAL_POLYNOMIAL_EVALUATE_FUNC(DATA_VECTOR(i, 1), ATEMP_VECTOR, 0)
    BTEMP_VAL = RATIONAL_POLYNOMIAL_EVALUATE_FUNC(DATA_VECTOR(i, 1), BTEMP_VECTOR, 0)
    If BTEMP_VAL <> 0 Then
        'numerator derivatives
        GRAD_VECTOR(i, 1) = 1 / BTEMP_VAL   'GRAD_VECTOR/da0
        For j = 2 To DEG_NUMER + 1
            GRAD_VECTOR(i, j) = GRAD_VECTOR(i, j - 1) * DATA_VECTOR(i, 1)
        Next j
        'denominator derivatives
        GRAD_VECTOR(i, DEG_NUMER + 2) = -ATEMP_VAL / BTEMP_VAL ^ 2
        'GRAD_VECTOR/db0
        For j = 2 To DEG_DENOM
            GRAD_VECTOR(i, j + DEG_NUMER + 1) = _
                GRAD_VECTOR(i, j + DEG_NUMER) * DATA_VECTOR(i, 1)
        Next j
    End If
Next i

RATIONAL_POLYNOMIAL_GRADIENT_FUNC = GRAD_VECTOR

Exit Function
ERROR_LABEL:
RATIONAL_POLYNOMIAL_GRADIENT_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : RATIONAL_POLYNOMIAL_EVALUATE_FUNC

'DESCRIPTION   : Returns the polynomial or its derivative at the point x
'P(x)=a0+a1*x+a2*x^2+...an*x^n
'where coeff(0)=a0, coeff(1)=a1, etc...

'LIBRARY       : POLYNOMIAL
'GROUP         : RATIONAL
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************


Function RATIONAL_POLYNOMIAL_EVALUATE_FUNC(ByVal X_VALUE As Double, _
ByRef COEF_RNG As Variant, _
Optional ByVal NSIZE As Long = 1)

Dim i As Long
Dim j As Long

Dim NROWS As Long

Dim Y_VAL As Double

Dim TEMP_VECTOR As Variant
Dim COEF_VECTOR As Variant
    
On Error GoTo ERROR_LABEL
    
COEF_VECTOR = COEF_RNG
NROWS = UBound(COEF_VECTOR, 1)

ReDim TEMP_VECTOR(0 To NROWS, 1 To 1)

If NSIZE > NROWS Then: GoTo ERROR_LABEL
'calcolo coefficienti derivata
For i = 0 To NROWS
    TEMP_VECTOR(i, 1) = COEF_VECTOR(NROWS - i, 1)
Next i

For i = 0 To NROWS - NSIZE
    For j = 1 To NSIZE
        TEMP_VECTOR(i, 1) = (NROWS - i - j + 1) * TEMP_VECTOR(i, 1)
    Next j
Next i
'calcolo derivata del polinomio
Y_VAL = 0
For i = 0 To NROWS - NSIZE
    Y_VAL = TEMP_VECTOR(i, 1) + Y_VAL * X_VALUE
Next i

RATIONAL_POLYNOMIAL_EVALUATE_FUNC = Y_VAL

Exit Function
ERROR_LABEL:
RATIONAL_POLYNOMIAL_EVALUATE_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : RATIONAL_POLYNOMIAL_NORMALIZE_FUNC
'DESCRIPTION   : Normalize coefficients of a rational polynomial
'LIBRARY       : POLYNOMIAL
'GROUP         : RATIONAL
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function RATIONAL_POLYNOMIAL_NORMALIZE_FUNC(ByRef COEF_RNG As Variant, _
Optional ByVal tolerance As Double = 10 ^ -15)

Dim i As Long
Dim j As Long
Dim SROW As Long
Dim NROWS As Long

Dim TEMP_FACTOR As Double
Dim COEF_VECTOR As Variant

On Error GoTo ERROR_LABEL

COEF_VECTOR = COEF_RNG
SROW = LBound(COEF_VECTOR, 1)
NROWS = UBound(COEF_VECTOR, 1)

For i = SROW To NROWS
    If Abs(COEF_VECTOR(i, 1)) > tolerance Then
        j = i
    Else
        COEF_VECTOR(i, 1) = 0
    End If
Next i

TEMP_FACTOR = COEF_VECTOR(j, 1)
For i = SROW To NROWS
    COEF_VECTOR(i, 1) = COEF_VECTOR(i, 1) / TEMP_FACTOR
Next i

RATIONAL_POLYNOMIAL_NORMALIZE_FUNC = COEF_VECTOR

Exit Function
ERROR_LABEL:
RATIONAL_POLYNOMIAL_NORMALIZE_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : RATIONAL_POLYNOMIAL_WRITE_FUNC
'DESCRIPTION   : Write Rational Polynomial String
'Rational Formula; numerator degree; denominator degree;
'"=(c1+c2*x+c3*x^2)/(c4+c5*x+c6*x^2+c7*x^3)"
'LIBRARY       : POLYNOMIAL
'GROUP         : RATIONAL
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function RATIONAL_POLYNOMIAL_WRITE_FUNC(ByVal NUMER_DEG As Long, _
ByVal DENOM_DEG As Long, _
Optional ByVal CONST_CHR As String = "c", _
Optional ByVal VAR_CHR As String = "x")

Dim i As Long
Dim j As Long
Dim FIRST_STR As String
Dim SECOND_STR As String

On Error GoTo ERROR_LABEL

i = 0
If NUMER_DEG = 0 Then
    i = i + 1
    FIRST_STR = CONST_CHR & i
Else
    i = i + 1
    FIRST_STR = CONST_CHR & i
    i = i + 1
    FIRST_STR = FIRST_STR & "+" & CONST_CHR & i & "*" & VAR_CHR
    For j = 2 To NUMER_DEG
        i = i + 1
        FIRST_STR = FIRST_STR & "+" & CONST_CHR & i & "*" & VAR_CHR & "^" & j
    Next j
End If

If DENOM_DEG = 0 Then
    i = i + 1
    SECOND_STR = CONST_CHR & i
Else
    i = i + 1
    SECOND_STR = CONST_CHR & i
    i = i + 1
    SECOND_STR = SECOND_STR & "+" & CONST_CHR & i & "*" & VAR_CHR
    For j = 2 To DENOM_DEG
        i = i + 1
        SECOND_STR = SECOND_STR & "+" & CONST_CHR & i & "*" & VAR_CHR & "^" & j
    Next j
End If

RATIONAL_POLYNOMIAL_WRITE_FUNC = "(" & FIRST_STR & ")/(" & SECOND_STR & ")"

Exit Function
ERROR_LABEL:
RATIONAL_POLYNOMIAL_WRITE_FUNC = Err.number
End Function
