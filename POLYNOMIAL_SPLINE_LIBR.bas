Attribute VB_Name = "POLYNOMIAL_SPLINE_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : POLYNOMIAL_SPLINE_INTERPOLATION_FUNC
'DESCRIPTION   : The Cubic Spline interpolation is based on fitting a
'cubic polynomial curve through all the given set of points, called knots,
'following these rules:

'1) the curve pass through all knots
'2) at each knot, the first and second derivatives of the two curves
'that meet there are equal
'e) at the two end-knots, the second derivatives of each curve equal
'zero (natural cubic spline constrains).

'Returns the Y at a given X on the natural cubic spline
'curve defined by a given set of points (knots).

'TARGET_VAL: one value containing the x value at which we want
'to compute the interpolation

'LIBRARY       : POLYNOMIAL
'GROUP         : SPLINE
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function POLYNOMIAL_SPLINE_INTERPOLATION_FUNC(ByRef XDATA_RNG As Variant, _
ByRef YDATA_RNG As Variant, _
ByVal TARGET_VAL As Double)

Dim i As Double
Dim j As Double
Dim k As Double

Dim NSIZE As Long
Dim NROWS As Long

Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double
Dim CTEMP_VAL As Double
Dim DTEMP_VAL As Double
Dim ETEMP_VAL As Double

Dim TEMP_VAL As Double
Dim MULT_VAL As Double

Dim SCALE0_VAL As Double
Dim SCALE1_VAL As Double
Dim SCALE2_VAL As Double

Dim TEMP_MATRIX As Variant

Dim XDATA_VECTOR As Variant
Dim YDATA_VECTOR As Variant

Dim ATEMP_VECTOR As Variant
Dim BTEMP_VECTOR As Variant

On Error GoTo ERROR_LABEL

XDATA_VECTOR = XDATA_RNG
If UBound(XDATA_VECTOR, 1) = 1 Then
    XDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(XDATA_VECTOR)
End If

YDATA_VECTOR = YDATA_RNG
If UBound(YDATA_VECTOR, 1) = 1 Then
    YDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(YDATA_VECTOR)
End If
NROWS = UBound(YDATA_VECTOR, 1)
NSIZE = UBound(XDATA_VECTOR, 1)

' Next check to be sure that there are the came counts
'of XDATA_RNG and YDATA_RNG
If NSIZE <> NROWS Then: GoTo ERROR_LABEL 'Uneven counts
'of XDATA_RNG and YDATA_RNG"
 
ReDim ATEMP_VECTOR(1 To NSIZE - 1, 1 To 1)
ReDim BTEMP_VECTOR(1 To NSIZE, 1 To 1)
'these are the 2nd derivative values

'populate and order the input arrays

ReDim TEMP_MATRIX(1 To NSIZE, 1 To 2)
For i = 1 To NSIZE
    TEMP_MATRIX(i, 1) = XDATA_VECTOR(i, 1)
    TEMP_MATRIX(i, 2) = YDATA_VECTOR(i, 1)
Next i
 
TEMP_MATRIX = MATRIX_QUICK_SORT_FUNC(TEMP_MATRIX, 1, 1)
 
For i = 1 To NSIZE
    XDATA_VECTOR(i, 1) = TEMP_MATRIX(i, 1)
    YDATA_VECTOR(i, 1) = TEMP_MATRIX(i, 2)
Next i

BTEMP_VECTOR(1, 1) = 0 'First knot boundary condition
ATEMP_VECTOR(1, 1) = 0   'First knot boundary condition

For i = 2 To NSIZE - 1
    SCALE0_VAL = XDATA_VECTOR(i + 0, 1) - XDATA_VECTOR(i - 1, 1)
    SCALE1_VAL = XDATA_VECTOR(i + 1, 1) - XDATA_VECTOR(i - 0, 1)
    SCALE2_VAL = XDATA_VECTOR(i + 1, 1) - XDATA_VECTOR(i - 1, 1)
    
    DTEMP_VAL = SCALE0_VAL / SCALE2_VAL
    CTEMP_VAL = DTEMP_VAL * BTEMP_VECTOR(i - 1, 1) + 2
    BTEMP_VECTOR(i, 1) = (DTEMP_VAL - 1) / CTEMP_VAL
    ATEMP_VECTOR(i, 1) = (YDATA_VECTOR(i + 1, 1) - YDATA_VECTOR(i, 1)) / SCALE1_VAL - (YDATA_VECTOR(i, 1) - YDATA_VECTOR(i - 1, 1)) / SCALE0_VAL
    ATEMP_VECTOR(i, 1) = (6 * ATEMP_VECTOR(i, 1) / SCALE2_VAL - DTEMP_VAL * ATEMP_VECTOR(i - 1, 1)) / CTEMP_VAL
Next i
    
MULT_VAL = 0
ETEMP_VAL = 0
BTEMP_VECTOR(NSIZE, 1) = (ETEMP_VAL - MULT_VAL * ATEMP_VECTOR(NSIZE - 1, 1)) / (MULT_VAL * BTEMP_VECTOR(NSIZE - 1, 1) + 1) 'Last knot boundary condition


For i = NSIZE - 1 To 1 Step -1
    BTEMP_VECTOR(i, 1) = BTEMP_VECTOR(i, 1) * BTEMP_VECTOR(i + 1, 1) + ATEMP_VECTOR(i, 1)
        'Backfill the 2nd derivatives
Next i
''''''''''''''''''''''''''''''''''''''''
'Spline evaluation at target XDATA_VECTOR point
'''''''''''''''''''''''''''''''''''''''''
'Find correct interval using halving binary search
i = 1
j = NSIZE
Do While (i < j - 1)
    k = (i + j) / 2   'Calculate the midpoint
    If (TARGET_VAL < XDATA_VECTOR(k, 1)) Then j = k Else i = k  'Narrow down the bounds
Loop
' i = beginning of the correct interval
SCALE0_VAL = XDATA_VECTOR(i + 1, 1) - XDATA_VECTOR(i, 1)
'Calc the width of the XDATA_VECTOR interval
ATEMP_VAL = (XDATA_VECTOR(i + 1, 1) - TARGET_VAL) / SCALE0_VAL
BTEMP_VAL = (TARGET_VAL - XDATA_VECTOR(i, 1)) / SCALE0_VAL
TEMP_VAL = ATEMP_VAL * YDATA_VECTOR(i, 1) + BTEMP_VAL * YDATA_VECTOR(i + 1, 1) + ((ATEMP_VAL ^ 3 - ATEMP_VAL) * BTEMP_VECTOR(i, 1) + (BTEMP_VAL ^ 3 - BTEMP_VAL) * BTEMP_VECTOR(i + 1, 1)) * (SCALE0_VAL ^ 2) / 6

POLYNOMIAL_SPLINE_INTERPOLATION_FUNC = TEMP_VAL

Exit Function
ERROR_LABEL:
POLYNOMIAL_SPLINE_INTERPOLATION_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : POLYNOMIAL_SPLINE_EVALUATE_FUNC

'DESCRIPTION   : Using the POLYNOMIAL_SPLINE_EVALUATE_FUNC function is faster than using
'POLYNOMIAL_SPLINE_INTERPOLATION_FUNC, because POLYNOMIAL_SPLINE_EVALUATE_FUNC uses the
'information of the 2nd derivatives and does not have to calculate them all over again like
'POLYNOMIAL_SPLINE_INTERPOLATION_FUNC

'Interpolates one point from sorted XDATA_VECTOR,YDATA_VECTOR
'data pairs using Cubic Spline Interpolation

'LIBRARY       : POLYNOMIAL
'GROUP         : SPLINE
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function POLYNOMIAL_SPLINE_EVALUATE_FUNC(ByRef XDATA_RNG As Variant, _
ByRef YDATA_RNG As Variant, _
ByRef DERIV_RNG As Variant, _
ByVal TARGET_VAL As Double)

Dim i As Double
Dim j As Double
Dim k As Double

Dim NSIZE As Long
Dim NROWS As Long

Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double

Dim TEMP_VAL As Double

Dim SCALE0_VAL As Double

Dim TEMP_MATRIX As Variant

Dim XDATA_VECTOR As Variant
Dim YDATA_VECTOR As Variant
Dim ZDATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

XDATA_VECTOR = XDATA_RNG
If UBound(XDATA_VECTOR, 1) = 1 Then
    XDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(XDATA_VECTOR)
End If

YDATA_VECTOR = YDATA_RNG
If UBound(YDATA_VECTOR, 1) = 1 Then
    YDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(YDATA_VECTOR)
End If

ZDATA_VECTOR = DERIV_RNG
If UBound(ZDATA_VECTOR, 1) = 1 Then
    ZDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(ZDATA_VECTOR)
End If

NROWS = UBound(YDATA_VECTOR, 1)
NSIZE = UBound(XDATA_VECTOR, 1)

' Next check to be sure that there are the came counts
'of XDATA_RNG and YDATA_RNG
If NSIZE <> NROWS Then: GoTo ERROR_LABEL 'Uneven counts
'of XDATA_RNG and YDATA_RNG"
 
'populate and order the input arrays

ReDim TEMP_MATRIX(1 To NSIZE, 1 To 2)
For i = 1 To NSIZE
    TEMP_MATRIX(i, 1) = XDATA_VECTOR(i, 1)
    TEMP_MATRIX(i, 2) = YDATA_VECTOR(i, 1)
Next i
 
TEMP_MATRIX = MATRIX_QUICK_SORT_FUNC(TEMP_MATRIX, 1, 1)
 
For i = 1 To NSIZE
    XDATA_VECTOR(i, 1) = TEMP_MATRIX(i, 1)
    YDATA_VECTOR(i, 1) = TEMP_MATRIX(i, 2)
Next i
 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Spline evaluation at target XDATA_VECTOR point
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Find correct interval using halving binary search
i = 1
j = NSIZE
Do While (i < j - 1)
    k = (i + j) / 2   'Calculate the midpoint
    If (TARGET_VAL < XDATA_VECTOR(k, 1)) Then j = k Else i = k  'Narrow down the bounds
Loop
' i = beginning of the correct interval
SCALE0_VAL = XDATA_VECTOR(i + 1, 1) - XDATA_VECTOR(i, 1)
'Calc the width of the XDATA_VECTOR interval
ATEMP_VAL = (XDATA_VECTOR(i + 1, 1) - TARGET_VAL) / SCALE0_VAL
BTEMP_VAL = (TARGET_VAL - XDATA_VECTOR(i, 1)) / SCALE0_VAL
TEMP_VAL = ATEMP_VAL * YDATA_VECTOR(i, 1) + BTEMP_VAL * YDATA_VECTOR(i + 1, 1) + ((ATEMP_VAL ^ 3 - ATEMP_VAL) * ZDATA_VECTOR(i, 1) + (BTEMP_VAL ^ 3 - BTEMP_VAL) * ZDATA_VECTOR(i + 1, 1)) * (SCALE0_VAL ^ 2) / 6

POLYNOMIAL_SPLINE_EVALUATE_FUNC = TEMP_VAL
Exit Function
ERROR_LABEL:
POLYNOMIAL_SPLINE_EVALUATE_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : POLYNOMIAL_SPLINE_COEFFICIENT_FUNC

'DESCRIPTION   : Find the cubic spline polynomials that fit the given knots

'Returns the coefficients of the cubic spline polynomials for
'a given set of points (knots). This function accepts also points
'[Xin, Yin] in any order (random knots)

'Returns an (n-1 x 4 ) array, where n is the number of knots. Each
'row contains the coefficients of the cubic polynomial fit for each
'segment [a3, a2, a1, a0]

'LIBRARY       : POLYNOMIAL
'GROUP         : SPLINE
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function POLYNOMIAL_SPLINE_COEFFICIENT_FUNC(ByRef XDATA_RNG As Variant, _
ByRef YDATA_RNG As Variant)

Dim i As Long

Dim NSIZE As Long
Dim NROWS As Long

Dim CTEMP_VAL As Double
Dim DTEMP_VAL As Double
Dim ETEMP_VAL As Double

Dim MULT_VAL As Double

Dim SCALE0_VAL As Double
Dim SCALE1_VAL As Double
Dim SCALE2_VAL As Double

Dim XDATA_VECTOR As Variant
Dim YDATA_VECTOR As Variant

Dim ATEMP_VECTOR As Variant
Dim BTEMP_VECTOR As Variant

Dim TEMP_MATRIX As Variant
Dim COEF_MATRIX As Variant

On Error GoTo ERROR_LABEL

XDATA_VECTOR = XDATA_RNG
If UBound(XDATA_VECTOR, 1) = 1 Then
    XDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(XDATA_VECTOR)
End If

YDATA_VECTOR = YDATA_RNG
If UBound(YDATA_VECTOR, 1) = 1 Then
    YDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(YDATA_VECTOR)
End If

NROWS = UBound(YDATA_VECTOR, 1)
NSIZE = UBound(XDATA_VECTOR, 1)

' Next check to be sure that there are the came counts
'of XDATA_RNG and YDATA_RNG
If NSIZE <> NROWS Then: GoTo ERROR_LABEL 'Uneven counts
'of XDATA_RNG and YDATA_RNG"
 
ReDim ATEMP_VECTOR(1 To NSIZE - 1, 1 To 1)
ReDim BTEMP_VECTOR(1 To NSIZE, 1 To 1)
'these are the 2nd derivative values
ReDim COEF_MATRIX(1 To NSIZE, 1 To 4)
'these are the cubic spline coefficients

'populate and order the input arrays

ReDim TEMP_MATRIX(1 To NSIZE, 1 To 2)
For i = 1 To NSIZE
    TEMP_MATRIX(i, 1) = XDATA_VECTOR(i, 1)
    TEMP_MATRIX(i, 2) = YDATA_VECTOR(i, 1)
Next i
 
TEMP_MATRIX = MATRIX_QUICK_SORT_FUNC(TEMP_MATRIX, 1, 1)
 
For i = 1 To NSIZE
    XDATA_VECTOR(i, 1) = TEMP_MATRIX(i, 1)
    YDATA_VECTOR(i, 1) = TEMP_MATRIX(i, 2)
Next i

BTEMP_VECTOR(1, 1) = 0 'First knot boundary condition
ATEMP_VECTOR(1, 1) = 0   'First knot boundary condition

For i = 2 To NSIZE - 1
    SCALE0_VAL = XDATA_VECTOR(i + 0, 1) - XDATA_VECTOR(i - 1, 1)
    SCALE1_VAL = XDATA_VECTOR(i + 1, 1) - XDATA_VECTOR(i - 0, 1)
    SCALE2_VAL = XDATA_VECTOR(i + 1, 1) - XDATA_VECTOR(i - 1, 1)
    
    DTEMP_VAL = SCALE0_VAL / SCALE2_VAL
    CTEMP_VAL = DTEMP_VAL * BTEMP_VECTOR(i - 1, 1) + 2
    BTEMP_VECTOR(i, 1) = (DTEMP_VAL - 1) / CTEMP_VAL
    
    ATEMP_VECTOR(i, 1) = (YDATA_VECTOR(i + 1, 1) - YDATA_VECTOR(i, 1)) / SCALE1_VAL - (YDATA_VECTOR(i, 1) - YDATA_VECTOR(i - 1, 1)) / SCALE0_VAL
    ATEMP_VECTOR(i, 1) = (6 * ATEMP_VECTOR(i, 1) / SCALE2_VAL - DTEMP_VAL * ATEMP_VECTOR(i - 1, 1)) / CTEMP_VAL

Next i
    
MULT_VAL = 0: ETEMP_VAL = 0
BTEMP_VECTOR(NSIZE, 1) = (ETEMP_VAL - MULT_VAL * ATEMP_VECTOR(NSIZE - 1, 1)) / (MULT_VAL * BTEMP_VECTOR(NSIZE - 1, 1) + 1) 'Last knot boundary condition

For i = NSIZE - 1 To 1 Step -1
    BTEMP_VECTOR(i, 1) = BTEMP_VECTOR(i, 1) * BTEMP_VECTOR(i + 1, 1) + ATEMP_VECTOR(i, 1) 'Backfill the 2nd derivatives
    SCALE0_VAL = XDATA_VECTOR(i + 1, 1) - XDATA_VECTOR(i, 1) 'Calculate the coefficients for each spline fragment
    COEF_MATRIX(i, 1) = (BTEMP_VECTOR(i + 1, 1) - BTEMP_VECTOR(i, 1)) / (6 * SCALE0_VAL) 'ATEMP (XDATA_VECTOR^3)
    COEF_MATRIX(i, 2) = BTEMP_VECTOR(i, 1) / 2 'BTEMP (XDATA_VECTOR^2)
    COEF_MATRIX(i, 3) = (YDATA_VECTOR(i + 1, 1) - YDATA_VECTOR(i, 1)) / SCALE0_VAL - (2 * SCALE0_VAL * BTEMP_VECTOR(i, 1) + SCALE0_VAL * BTEMP_VECTOR(i + 1, 1)) / 6 'C (XDATA_VECTOR)
    COEF_MATRIX(i, 4) = YDATA_VECTOR(i, 1) 'D
Next i

POLYNOMIAL_SPLINE_COEFFICIENT_FUNC = COEF_MATRIX

Exit Function
ERROR_LABEL:
POLYNOMIAL_SPLINE_COEFFICIENT_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : POLYNOMIAL_SPLINE_GRADIENT_FUNC

'DESCRIPTION   : Returns the cubic spline 2nd derivatives at a given set
'of points (knots).

'This function accepts also points [Xin, Yin] in any order (random knots)
'For n knots , returns an n array of the 2nd derivative values. The first
'and the last values are zero (natural spline constrain)
'We note that the 2nd derivatives depend only by the given set of knots.
'So this function can be evaluate only once for the whole range of
'interpolation.

'LIBRARY       : POLYNOMIAL
'GROUP         : SPLINE
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function POLYNOMIAL_SPLINE_GRADIENT_FUNC(ByRef XDATA_RNG As Variant, _
ByRef YDATA_RNG As Variant)

'The natural cubic spline has a continuous second derivative
'(acceleration). This characteristic is very important in many
'applied sciences (Numeric Control, Automation, etc..) when we
'need to reduce the vibration and the noise in electromechanical
'motions, although the cubic spline is much slower than other
'interpolation methods.

Dim i As Long
Dim NSIZE As Long
Dim NROWS As Long

Dim CTEMP_VAL As Double
Dim DTEMP_VAL As Double
Dim ETEMP_VAL As Double

Dim MULT_VAL As Double

Dim SCALE0_VAL As Double
Dim SCALE1_VAL As Double
Dim SCALE2_VAL As Double

Dim TEMP_MATRIX As Variant
Dim ATEMP_VECTOR As Variant
Dim BTEMP_VECTOR As Variant

Dim XDATA_VECTOR As Variant
Dim YDATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

XDATA_VECTOR = XDATA_RNG
If UBound(XDATA_VECTOR, 1) = 1 Then
    XDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(XDATA_VECTOR)
End If

YDATA_VECTOR = YDATA_RNG
If UBound(YDATA_VECTOR, 1) = 1 Then
    YDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(YDATA_VECTOR)
End If

NROWS = UBound(YDATA_VECTOR, 1)
NSIZE = UBound(XDATA_VECTOR, 1)

' Next check to be sure that there are the came counts
'of XDATA_RNG and YDATA_RNG
If NSIZE <> NROWS Then: GoTo ERROR_LABEL 'Uneven counts
'of XDATA_RNG and YDATA_RNG"
 
ReDim ATEMP_VECTOR(1 To NSIZE - 1, 1 To 1)
ReDim BTEMP_VECTOR(1 To NSIZE, 1 To 1)
'these are the 2nd derivative values

'populate and order the input arrays

ReDim TEMP_MATRIX(1 To NSIZE, 1 To 2)
For i = 1 To NSIZE
    TEMP_MATRIX(i, 1) = XDATA_VECTOR(i, 1)
    TEMP_MATRIX(i, 2) = YDATA_VECTOR(i, 1)
Next i
 
TEMP_MATRIX = MATRIX_QUICK_SORT_FUNC(TEMP_MATRIX, 1, 1)
 
For i = 1 To NSIZE
    XDATA_VECTOR(i, 1) = TEMP_MATRIX(i, 1)
    YDATA_VECTOR(i, 1) = TEMP_MATRIX(i, 2)
Next i

BTEMP_VECTOR(1, 1) = 0 'First knot boundary condition
ATEMP_VECTOR(1, 1) = 0   'First knot boundary condition

For i = 2 To NSIZE - 1
    SCALE0_VAL = XDATA_VECTOR(i + 0, 1) - XDATA_VECTOR(i - 1, 1)
    SCALE1_VAL = XDATA_VECTOR(i + 1, 1) - XDATA_VECTOR(i - 0, 1)
    SCALE2_VAL = XDATA_VECTOR(i + 1, 1) - XDATA_VECTOR(i - 1, 1)
    DTEMP_VAL = SCALE0_VAL / SCALE2_VAL
    CTEMP_VAL = DTEMP_VAL * BTEMP_VECTOR(i - 1, 1) + 2
    BTEMP_VECTOR(i, 1) = (DTEMP_VAL - 1) / CTEMP_VAL
    ATEMP_VECTOR(i, 1) = (YDATA_VECTOR(i + 1, 1) - YDATA_VECTOR(i, 1)) / SCALE1_VAL - (YDATA_VECTOR(i, 1) - YDATA_VECTOR(i - 1, 1)) / SCALE0_VAL
    ATEMP_VECTOR(i, 1) = (6 * ATEMP_VECTOR(i, 1) / SCALE2_VAL - DTEMP_VAL * ATEMP_VECTOR(i - 1, 1)) / CTEMP_VAL
Next i
    
MULT_VAL = 0
ETEMP_VAL = 0
BTEMP_VECTOR(NSIZE, 1) = (ETEMP_VAL - MULT_VAL * ATEMP_VECTOR(NSIZE - 1, 1)) / (MULT_VAL * BTEMP_VECTOR(NSIZE - 1, 1) + 1) _
'Last knot boundary condition
For i = NSIZE - 1 To 1 Step -1
    BTEMP_VECTOR(i, 1) = BTEMP_VECTOR(i, 1) * BTEMP_VECTOR(i + 1, 1) + ATEMP_VECTOR(i, 1) _
    'Backfill the 2nd derivatives
Next i

POLYNOMIAL_SPLINE_GRADIENT_FUNC = BTEMP_VECTOR

Exit Function
ERROR_LABEL:
POLYNOMIAL_SPLINE_GRADIENT_FUNC = Err.number
End Function
