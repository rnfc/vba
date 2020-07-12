Attribute VB_Name = "STAT_REGRESSION_RESID_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.



'************************************************************************************
'************************************************************************************
'FUNCTION      : RESIDUALS_REGRESSION_FUNC
'DESCRIPTION   : RETURNS THE RESIDUALS ON A MULTIPLE REGRESSION ANALYSIS
'LIBRARY       : STATISTICS
'GROUP         : RESIDUAL
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function RESIDUALS_REGRESSION_FUNC(ByRef XDATA_RNG As Variant, _
ByRef YDATA_RNG As Variant, _
Optional ByVal INTERCEPT_FLAG As Boolean = True, _
Optional ByVal OUTPUT As Integer = 0)
    
Dim i As Long
Dim j As Long
Dim k As Long

Dim NROWS As Long
Dim NCOLUMNS As Long
    
Dim TEMP_MATRIX As Variant

Dim YDATA_VECTOR As Variant
Dim XDATA_MATRIX As Variant

Dim SSR_VECTOR As Variant
Dim COEF_VECTOR As Variant
Dim RESIDUAL_MATRIX As Variant

Dim TEMP_SUM As Double

On Error GoTo ERROR_LABEL

XDATA_MATRIX = XDATA_RNG
YDATA_VECTOR = YDATA_RNG

NROWS = UBound(XDATA_MATRIX, 1)
NCOLUMNS = UBound(XDATA_MATRIX, 2)
    
If UBound(XDATA_MATRIX, 1) <> UBound(YDATA_VECTOR, 1) Then: GoTo ERROR_LABEL

ReDim SSR_VECTOR(1 To 1, 1 To NCOLUMNS)
ReDim RESIDUAL_MATRIX(1 To NROWS, 1 To NCOLUMNS)

ReDim TEMP_MATRIX(1 To NROWS, 1 To NCOLUMNS)

'-------------------------------------------------------------------------------------------------
For k = 1 To NCOLUMNS ' First get the SLOPE and ALPHA of Rand regressed on X
'-------------------------------------------------------------------------------------------------
    ReDim Preserve TEMP_MATRIX(1 To NROWS, 1 To k)
    ReDim Preserve RESIDUAL_MATRIX(1 To NROWS, 1 To k)
    For i = 1 To NROWS
        For j = 1 To k
            TEMP_MATRIX(i, j) = XDATA_MATRIX(i, j)
        Next j
    Next i
    COEF_VECTOR = REGRESSION_MULT_COEF_FUNC(TEMP_MATRIX, YDATA_VECTOR, INTERCEPT_FLAG, 0)
'-------------------------------------------------------------------------------------------------
    If INTERCEPT_FLAG = True Then
'-------------------------------------------------------------------------------------------------
       For i = 1 To NROWS
            TEMP_SUM = COEF_VECTOR(1, 1)
            For j = 1 To k
                TEMP_SUM = TEMP_SUM + COEF_VECTOR(j + 1, 1) * TEMP_MATRIX(i, j)
                ' Compute the residuals
            Next j
            RESIDUAL_MATRIX(i, k) = YDATA_VECTOR(i, 1) - TEMP_SUM
            SSR_VECTOR(1, k) = SSR_VECTOR(1, k) + RESIDUAL_MATRIX(i, k) ^ 2
       Next i
'-------------------------------------------------------------------------------------------------
    ElseIf INTERCEPT_FLAG = False Then
'-------------------------------------------------------------------------------------------------
        For i = 1 To NROWS
            TEMP_SUM = 0
            For j = 1 To k
                TEMP_SUM = TEMP_SUM + COEF_VECTOR(j, 1) * TEMP_MATRIX(i, j)
                ' Compute the residuals
            Next j
            RESIDUAL_MATRIX(i, k) = YDATA_VECTOR(i, 1) - TEMP_SUM
            SSR_VECTOR(1, k) = SSR_VECTOR(1, k) + RESIDUAL_MATRIX(i, k) ^ 2
        Next i
'-------------------------------------------------------------------------------------------------
    End If
'-------------------------------------------------------------------------------------------------
Next k
'-------------------------------------------------------------------------------------------------
    
Select Case OUTPUT
Case 0
    RESIDUALS_REGRESSION_FUNC = RESIDUAL_MATRIX
Case 1
    RESIDUALS_REGRESSION_FUNC = SSR_VECTOR
End Select

Exit Function
ERROR_LABEL:
RESIDUALS_REGRESSION_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : RESIDUAL_SCALE_FUNC
'DESCRIPTION   : RE-SCALE RESIDUAL MATRIX
'LIBRARY       : STATISTICS
'GROUP         : RESIDUAL
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function RESIDUAL_SCALE_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal VERSION As Integer = 0)

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim MEAN_VECTOR As Variant
Dim SIGMA_VECTOR As Variant
Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

'-----------------------------------------------------------------------------------
Select Case VERSION
'-----------------------------------------------------------------------------------
Case 0
    SIGMA_VECTOR = MATRIX_STDEV_FUNC(DATA_MATRIX)
    ReDim TEMP_MATRIX(1 To NROWS, 1 To NCOLUMNS)
    For j = 1 To NCOLUMNS
        For i = 1 To NROWS
            TEMP_MATRIX(i, j) = (1 / SIGMA_VECTOR(1, j)) * DATA_MATRIX(i, j) + 1
        Next i
    Next j
'-----------------------------------------------------------------------------------
Case Else
'-----------------------------------------------------------------------------------
    MEAN_VECTOR = MATRIX_MEAN_FUNC(DATA_MATRIX)
    SIGMA_VECTOR = MATRIX_STDEV_FUNC(DATA_MATRIX)
    ReDim TEMP_MATRIX(1 To NROWS, 1 To NCOLUMNS)
    For j = 1 To NCOLUMNS
        For i = 1 To NROWS
            TEMP_MATRIX(i, j) = (1 / SIGMA_VECTOR(1, j)) * (DATA_MATRIX(i, j) - MEAN_VECTOR(1, j)) + 1
        Next i
    Next j
'-----------------------------------------------------------------------------------
End Select
'-----------------------------------------------------------------------------------

RESIDUAL_SCALE_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
RESIDUAL_SCALE_FUNC = Err.number
End Function


Function REGRESSION_TARGET_DEPENDENT_VECTOR_FUNC(ByVal MU_VAL As Double, _
ByVal SE_VAL As Double, _
ByVal RHO_VAL As Double, _
ByRef XDATA_RNG As Variant, _
ByRef YDATA_RNG As Variant, _
Optional ByVal INTERCEPT_FLAG As Boolean = True)

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim SCALAR1_VAL As Double
Dim SCALAR2_VAL As Double

Dim XDATA_MATRIX As Variant
Dim YDATA_VECTOR As Variant

Dim XMEAN_VECTOR As Variant
Dim XSIGMA_VECTOR As Variant
Dim YTEMP_VECTOR As Variant

Dim SCALE_VECTOR As Variant
Dim RATIO_VECTOR As Variant
Dim RSIGMA_VECTOR As Variant

Dim TEMP_MATRIX As Variant
Dim RESIDUAL_MATRIX As Variant

On Error GoTo ERROR_LABEL

XDATA_MATRIX = XDATA_RNG
YDATA_VECTOR = YDATA_RNG

If UBound(XDATA_MATRIX, 1) <> UBound(YDATA_VECTOR, 1) Then: GoTo ERROR_LABEL
If (RHO_VAL > 1) Or (RHO_VAL < -1) Then: GoTo ERROR_LABEL
'Rho must be a number between -1 and 1 inclusive

NROWS = UBound(XDATA_MATRIX, 1)
NCOLUMNS = UBound(XDATA_MATRIX, 2)
        
RESIDUAL_MATRIX = RESIDUALS_REGRESSION_FUNC(XDATA_MATRIX, YDATA_VECTOR, INTERCEPT_FLAG, 0)
    
XSIGMA_VECTOR = MATRIX_STDEV_FUNC(XDATA_MATRIX)
RSIGMA_VECTOR = MATRIX_STDEV_FUNC(RESIDUAL_MATRIX)
XMEAN_VECTOR = MATRIX_MEAN_FUNC(XDATA_MATRIX)
        
ReDim RATIO_VECTOR(1 To 1, 1 To NCOLUMNS)
For j = 1 To NCOLUMNS
    RATIO_VECTOR(1, j) = XSIGMA_VECTOR(1, j) / RSIGMA_VECTOR(1, j)
Next j
     
Select Case RHO_VAL
Case -1
    SCALAR1_VAL = -1
    SCALAR2_VAL = Sqr((1 / (-RHO_VAL) ^ 2) - 1)
Case Is < 0 And RHO_VAL > -1
    SCALAR1_VAL = -1
    SCALAR2_VAL = Sqr((1 / (-RHO_VAL) ^ 2) - 1)
Case 0
    SCALAR1_VAL = 0
    SCALAR2_VAL = 1
Case Is < 1 And RHO_VAL > 0
    SCALAR1_VAL = 1
    SCALAR2_VAL = Sqr((1 / RHO_VAL ^ 2) - 1)
Case 1
    SCALAR1_VAL = 1
    SCALAR2_VAL = Sqr((1 / RHO_VAL ^ 2) - 1)
Case Else
    GoTo ERROR_LABEL
End Select

ReDim YTEMP_VECTOR(1 To 1, 1 To NCOLUMNS)
ReDim SCALE_VECTOR(1 To 1, 1 To NCOLUMNS)
ReDim TEMP_MATRIX(1 To NROWS, 1 To NCOLUMNS)
    
' Get the E_Y constant
For j = 1 To NCOLUMNS
    If Abs(RHO_VAL) <> 1 Then
        YTEMP_VECTOR(1, j) = (1 / SCALAR2_VAL) * (MU_VAL - SCALAR1_VAL * XMEAN_VECTOR(1, j))
    Else
        YTEMP_VECTOR(1, j) = MU_VAL - SCALAR1_VAL * XMEAN_VECTOR(1, j)
    End If
        
    For i = 1 To NROWS
        TEMP_MATRIX(i, j) = SCALAR1_VAL * XDATA_MATRIX(i, j) + SCALAR2_VAL * (RATIO_VECTOR(1, j) * RESIDUAL_MATRIX(i, j) + YTEMP_VECTOR(1, j))
    Next i
Next j

XSIGMA_VECTOR = MATRIX_STDEVP_FUNC(TEMP_MATRIX)
For j = 1 To NCOLUMNS
    SCALE_VECTOR(1, j) = SE_VAL / XSIGMA_VECTOR(1, j)
    For i = 1 To NROWS
        TEMP_MATRIX(i, j) = SCALE_VECTOR(1, j) * TEMP_MATRIX(i, j) + (MU_VAL * (1 - SCALE_VECTOR(1, j))) '
    Next i
Next j

REGRESSION_TARGET_DEPENDENT_VECTOR_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
REGRESSION_TARGET_DEPENDENT_VECTOR_FUNC = Err.number
End Function


