Attribute VB_Name = "STAT_REGRESSION_CIRC_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
        'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : CIRCULAR_REGRESSION_FUNC

'DESCRIPTION   : It computes the LS circular regression of a dataset [xi, yi].
'It returns a (3 x 2) array containing the radius R and the
'center coordinate (Xc, Yc) of the best fitting circle in the
'sense of the least squares. The second column contains the
'standard deviations of the estimates.

'(x, y) = data set to fit
' r = radius
'(xc, yc) = center
'(sxc, syc) = center standard deviations
' sr = radius standard deviation

'LIBRARY       : STATISTICS
'GROUP         : REGRESSION_CIRCULAR
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 11/29/2008
'************************************************************************************
'************************************************************************************

Function CIRCULAR_REGRESSION_FUNC(ByRef XDATA_RNG As Variant, _
ByRef YDATA_RNG As Variant, _
Optional ByVal VERSION As Integer = 0)

Dim i As Long
Dim NROWS As Long

Dim XC_VAL As Double
Dim YC_VAL As Double
Dim SR_VAL As Double
Dim SX_VAL As Double
Dim SY_VAL As Double
Dim RAD_VAL As Double

Dim XTEMP_SUM As Double
Dim YTEMP_SUM As Double

Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

Dim XDATA_VECTOR As Variant
Dim YDATA_VECTOR As Variant

Dim XTEMP_VECTOR As Variant
Dim YTEMP_VECTOR As Variant

Dim tolerance As Double

On Error GoTo ERROR_LABEL

tolerance = 10 ^ -21

XDATA_VECTOR = XDATA_RNG
If UBound(XDATA_VECTOR, 1) = 1 Then: XDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(XDATA_VECTOR)

YDATA_VECTOR = YDATA_RNG
If UBound(YDATA_VECTOR, 1) = 1 Then: YDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(YDATA_VECTOR)

NROWS = UBound(XDATA_VECTOR, 1)

'-----------------------------------------------------------------------------------
Select Case VERSION
'-----------------------------------------------------------------------------------
Case 0
'-----------------------------------------------------------------------------------
    ReDim TEMP_MATRIX(1 To 3, 1 To 2)
    ReDim XTEMP_VECTOR(1 To NROWS, 1 To 2)
    ReDim YTEMP_VECTOR(1 To NROWS, 1 To 1)
    
    For i = 1 To NROWS
        XTEMP_VECTOR(i, 1) = XDATA_VECTOR(i, 1)
        XTEMP_VECTOR(i, 2) = YDATA_VECTOR(i, 1)
    Next i
    
    For i = 1 To NROWS 'set the constant vector
        YTEMP_VECTOR(i, 1) = -(XTEMP_VECTOR(i, 1) ^ 2 + XTEMP_VECTOR(i, 2) ^ 2)
    Next i
    
    DATA_MATRIX = MATRIX_SVD_REGRESSION_FUNC(XTEMP_VECTOR, YTEMP_VECTOR, True) 'radius
    RAD_VAL = (DATA_MATRIX(2, 1) / 2) ^ 2 + (DATA_MATRIX(3, 1) / 2) ^ 2 - DATA_MATRIX(1, 1)
               
    If RAD_VAL <= tolerance Then GoTo ERROR_LABEL
    
    RAD_VAL = Sqr(RAD_VAL) 'center
    XC_VAL = -DATA_MATRIX(2, 1) / 2
    YC_VAL = -DATA_MATRIX(3, 1) / 2 'standard deviations
        
    SX_VAL = DATA_MATRIX(2, 2) / 2
    SY_VAL = DATA_MATRIX(3, 2) / 2
    
    SR_VAL = Sqr((XC_VAL * SX_VAL) ^ 2 + (YC_VAL * SY_VAL) ^ 2 + _
    DATA_MATRIX(1, 2) ^ 2 / 4) / RAD_VAL 'put togheter
    
    '-------------------PARAMETERS-----------------
    TEMP_MATRIX(1, 1) = RAD_VAL
    TEMP_MATRIX(2, 1) = XC_VAL
    TEMP_MATRIX(3, 1) = YC_VAL
    
    '---------------------SIGMA--------------------
    TEMP_MATRIX(1, 2) = SR_VAL
    TEMP_MATRIX(2, 2) = SX_VAL
    TEMP_MATRIX(3, 2) = SY_VAL
    '-----------------------------------------------
'-----------------------------------------------------------------------------------
Case Else
'-----------------------------------------------------------------------------------

    ReDim TEMP_MATRIX(1 To 3, 1 To 1)
    ReDim XTEMP_VECTOR(1 To NROWS, 1 To 2)
    ReDim YTEMP_VECTOR(1 To NROWS, 1 To 1)
        
    XTEMP_SUM = 0
    YTEMP_SUM = 0
    For i = 1 To NROWS 'compute the average of x and y
        XTEMP_SUM = XTEMP_SUM + XDATA_VECTOR(i, 1)
        YTEMP_SUM = YTEMP_SUM + YDATA_VECTOR(i, 1)
    Next i
    
    XTEMP_SUM = XTEMP_SUM / NROWS
    YTEMP_SUM = YTEMP_SUM / NROWS
    
    For i = 1 To NROWS 'shift the data
        XTEMP_VECTOR(i, 1) = XDATA_VECTOR(i, 1) - XTEMP_SUM
        XTEMP_VECTOR(i, 2) = YDATA_VECTOR(i, 1) - YTEMP_SUM
    Next i
    
    For i = 1 To NROWS 'set the constant vector
        YTEMP_VECTOR(i, 1) = -(XTEMP_VECTOR(i, 1) ^ 2 + XTEMP_VECTOR(i, 2) ^ 2)
    Next i
    DATA_MATRIX = MATRIX_SVD_REGRESSION_FUNC(XTEMP_VECTOR, YTEMP_VECTOR, True) 'radius
    
    RAD_VAL = (DATA_MATRIX(1, 1) / 2) ^ 2 + (DATA_MATRIX(2, 1) / 2) ^ 2 - DATA_MATRIX(3, 1)
    
    If RAD_VAL <= tolerance Then: GoTo ERROR_LABEL 'circle not found
    
    'compute radius and center
    RAD_VAL = Sqr(RAD_VAL)
    XC_VAL = -DATA_MATRIX(1, 1) / 2 + XTEMP_SUM
    YC_VAL = -DATA_MATRIX(2, 1) / 2 + YTEMP_SUM
            
    TEMP_MATRIX(1, 1) = RAD_VAL
    TEMP_MATRIX(2, 1) = XC_VAL
    TEMP_MATRIX(3, 1) = YC_VAL
'-----------------------------------------------------------------------------------
End Select
'-----------------------------------------------------------------------------------

CIRCULAR_REGRESSION_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
CIRCULAR_REGRESSION_FUNC = Err.number
End Function
