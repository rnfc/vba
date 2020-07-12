Attribute VB_Name = "STAT_REGRESSION_ORTH_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : ORTHOGONAL_DISTANCE_FUNC
'DESCRIPTION   : Orthogonal Distance
'LIBRARY       : STATISTICS
'GROUP         : REGRESSION_ORTHOGONAL
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/21/2009
'************************************************************************************
'************************************************************************************

Function ORTHOGONAL_DISTANCE_FUNC(ByRef XDATA_RNG As Variant, _
ByRef YDATA_RNG As Variant, _
ByVal X0_VAL As Double, _
ByVal Y0_VAL As Double)

'In ordinary linear regression, the goal is to minimize the sum of the
'squared vertical distances between the y data values and the corresponding
'y values on the fitted line. In orthogonal regression the goal is to minimize
'the orthogonal (perpendicular) distances from the data points to the
'fitted line.

'The slope-intercept equation for a line is:
'   Y = m*X + b
'where m is the slope and b is the intercept.

'A line perpendicular to this line will have -(1/m) slope, so the equation
'will be: Y' = -X/m + b'

'If this line passes through some data point (X0,Y0), its equation will be:
'Y' = -X/m + (X0/m + Y0)

'The perpendicular line will intersect the fitted line at a point (Xi,Yi)
'where,
'Xi and Yi are defined by:
'Xi = (X0 + m*Y0 - m*b) / (m^2 + 1)
'Yi = m*Xi + b

'So the orthogonal distance from (X0,X0) to the fitted line is the distance
'between (X0,Y0) and (Xi,Yi) which is computed as:
'distance = sqrt((X0-Xi)^2 + (Y0-Yi)^2)

Dim X1_VAL As Double
Dim Y1_VAL As Double

Dim SLOPE_VAL As Double
Dim ALPHA_VAL As Double

Dim XDATA_VECTOR As Variant
Dim YDATA_VECTOR As Variant

Dim TEMP_VECTOR As Variant
    
On Error GoTo ERROR_LABEL

XDATA_VECTOR = XDATA_RNG
YDATA_VECTOR = YDATA_RNG
TEMP_VECTOR = ORTHOGONAL_REGRESSION_FUNC(XDATA_VECTOR, YDATA_VECTOR)

SLOPE_VAL = TEMP_VECTOR(1, 1)
ALPHA_VAL = TEMP_VECTOR(2, 1)

X1_VAL = (X0_VAL + SLOPE_VAL * Y0_VAL - SLOPE_VAL * ALPHA_VAL) / (SLOPE_VAL ^ 2 + 1)
Y1_VAL = SLOPE_VAL * X1_VAL + ALPHA_VAL
    
ORTHOGONAL_DISTANCE_FUNC = Sqr((X0_VAL - X1_VAL) ^ 2 + (Y0_VAL - Y1_VAL) ^ 2)

Exit Function
ERROR_LABEL:
ORTHOGONAL_DISTANCE_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : ORTHOGONAL_REGRESSION_FUNC
'DESCRIPTION   : Least-squares Orthogonal Regression
'LIBRARY       : STATISTICS
'GROUP         : REGRESSION_ORTHOGONAL
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/21/2009
'************************************************************************************
'************************************************************************************

Function ORTHOGONAL_REGRESSION_FUNC(ByRef XDATA_RNG As Variant, _
ByRef YDATA_RNG As Variant)
    
Dim i As Long
Dim NROWS As Long

Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double

Dim V2_SUM_VAL As Double
Dim U2_SUM_VAL As Double
Dim UV_SUM_VAL As Double

Dim TRATIO_VAL As Double

Dim UTEMP_VECTOR As Variant
Dim VTEMP_VECTOR As Variant
Dim WTEMP_VECTOR As Variant

Dim XDATA_VECTOR As Variant
Dim YDATA_VECTOR As Variant

Dim XMEAN_VECTOR As Variant
Dim YMEAN_VECTOR As Variant

Dim TEMP_MATRIX As Variant

Dim XVOLATILITY_VECTOR As Variant
Dim YVOLATILITY_VECTOR As Variant

On Error GoTo ERROR_LABEL

XDATA_VECTOR = XDATA_RNG
If UBound(XDATA_VECTOR, 1) = 1 Then
    XDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(XDATA_VECTOR)
End If
NROWS = UBound(XDATA_VECTOR, 1)

YDATA_VECTOR = YDATA_RNG
If UBound(YDATA_VECTOR, 1) = 1 Then
    YDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(YDATA_VECTOR)
End If
        
ReDim UTEMP_VECTOR(1 To NROWS, 1 To 1)
ReDim VTEMP_VECTOR(1 To NROWS, 1 To 1)
ReDim WTEMP_VECTOR(1 To 2, 1 To 1)

XMEAN_VECTOR = MATRIX_MEAN_FUNC(XDATA_VECTOR)
YMEAN_VECTOR = MATRIX_MEAN_FUNC(YDATA_VECTOR)

XVOLATILITY_VECTOR = MATRIX_STDEV_FUNC(XDATA_VECTOR)
YVOLATILITY_VECTOR = MATRIX_STDEV_FUNC(YDATA_VECTOR)

For i = 1 To NROWS
    UTEMP_VECTOR(i, 1) = XDATA_VECTOR(i, 1) - XMEAN_VECTOR(1, 1)
    VTEMP_VECTOR(i, 1) = YDATA_VECTOR(i, 1) - YMEAN_VECTOR(1, 1)
Next i

V2_SUM_VAL = MATRIX_ELEMENTS_CUMULATIVE_SUM_FUNC(VECTOR_ELEMENTS_MULT_FUNC(VTEMP_VECTOR, VTEMP_VECTOR))
U2_SUM_VAL = MATRIX_ELEMENTS_CUMULATIVE_SUM_FUNC(VECTOR_ELEMENTS_MULT_FUNC(UTEMP_VECTOR, UTEMP_VECTOR))
UV_SUM_VAL = MATRIX_ELEMENTS_CUMULATIVE_SUM_FUNC(VECTOR_ELEMENTS_MULT_FUNC(UTEMP_VECTOR, VTEMP_VECTOR))

ATEMP_VAL = V2_SUM_VAL - U2_SUM_VAL
BTEMP_VAL = Sqr((U2_SUM_VAL - V2_SUM_VAL) ^ 2 + 4 * UV_SUM_VAL * UV_SUM_VAL)

WTEMP_VECTOR(1, 1) = (ATEMP_VAL + BTEMP_VAL) / (2 * UV_SUM_VAL)
WTEMP_VECTOR(2, 1) = (ATEMP_VAL - BTEMP_VAL) / (2 * UV_SUM_VAL)
TRATIO_VAL = UV_SUM_VAL / Sqr(U2_SUM_VAL * V2_SUM_VAL)

TEMP_MATRIX = REGRESSION_SIMPLE_COEF_FUNC(XDATA_VECTOR, YDATA_VECTOR)
If Sgn(WTEMP_VECTOR(1, 1)) = Sgn(TEMP_MATRIX(1, 1)) Then 'Slope
    TEMP_MATRIX(1, 1) = WTEMP_VECTOR(1, 1)
Else
    TEMP_MATRIX(1, 1) = WTEMP_VECTOR(2, 1)
End If

TEMP_MATRIX(2, 1) = YMEAN_VECTOR(1, 1) - TEMP_MATRIX(1, 1) * XMEAN_VECTOR(1, 1) 'intercept
TEMP_MATRIX(3, 1) = TEMP_MATRIX(1, 1) * Sqr((1 - TRATIO_VAL * TRATIO_VAL) / NROWS) / TRATIO_VAL 'sigma_slope

TEMP_MATRIX(4, 1) = _
    Sqr(((YVOLATILITY_VECTOR(1, 1) - XVOLATILITY_VECTOR(1, 1) * _
    TEMP_MATRIX(1, 1)) ^ 2) / NROWS + (1 - TRATIO_VAL) * _
    TEMP_MATRIX(1, 1) * (2 * XVOLATILITY_VECTOR(1, 1) * _
    YVOLATILITY_VECTOR(1, 1) + (XMEAN_VECTOR(1, 1) * _
    TEMP_MATRIX(1, 1) * (1 + TRATIO_VAL) / _
    (TRATIO_VAL * TRATIO_VAL))))
    'sigma_intercept

ORTHOGONAL_REGRESSION_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
ORTHOGONAL_REGRESSION_FUNC = Err.number
End Function
