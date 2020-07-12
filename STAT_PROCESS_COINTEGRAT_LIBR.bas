Attribute VB_Name = "STAT_PROCESS_COINTEGRAT_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : COINTEGRATION_EG_TEST_FUNC
'DESCRIPTION   : This function tests if the input series are I(1), and if the
'cointegration regression residuals are I(0), according to the Engle-Granger
'methodology.

'LIBRARY       : STATISTICS
'GROUP         : COINTEGRATION
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function COINTEGRATION_EG_TEST_FUNC(ByRef XDATA_RNG As Variant, _
Optional ByRef YDATA_RNG As Variant, _
Optional ByVal nLAGS As Long = 4)

Dim i As Long
Dim j As Long
Dim k As Long
Dim NROWS As Long
Dim NCOLUMNS As Long

Dim XDATA_MATRIX As Variant
Dim YDATA_VECTOR As Variant

Dim XTEMP_VECTOR As Variant
Dim YTEMP_VECTOR As Variant

Dim TEMP_MATRIX As Variant
Dim RESIDUAL_MATRIX As Variant

On Error GoTo ERROR_LABEL

XDATA_MATRIX = XDATA_RNG
NROWS = UBound(XDATA_MATRIX, 1)
NCOLUMNS = UBound(XDATA_MATRIX, 2)
'-------------------------------------------------------------------------------
If IsArray(YDATA_RNG) = False Then
'-------------------------------------------------------------------------------

    ReDim TEMP_MATRIX(0 To nLAGS + 1, 1 To NCOLUMNS + 1)
    TEMP_MATRIX(0, 1) = "LAGS"
    
    For k = 0 To nLAGS
        TEMP_MATRIX(k + 1, 1) = k
        For j = 1 To NCOLUMNS
            TEMP_MATRIX(0, j + 1) = "VAR " & CStr(j)
            ReDim XTEMP_VECTOR(1 To NROWS - k, 1 To 1)
            ReDim YTEMP_VECTOR(1 To NROWS - (k + 1), 1 To 1)
            For i = 1 To NROWS - k
                If k = 0 Then
                    XTEMP_VECTOR(i, 1) = XDATA_MATRIX(i, j)
                Else
                    XTEMP_VECTOR(i, 1) = XDATA_MATRIX(i + k, j) - XDATA_MATRIX(i + (k - 1), j)
                End If
            Next i
            For i = 1 To NROWS - (k + 1)
                YTEMP_VECTOR(i, 1) = XTEMP_VECTOR(i + 1, 1) - XTEMP_VECTOR(i, 1)
            Next i
            TEMP_MATRIX(k + 1, j + 1) = COINTEGRATION_EG_STATISTICS_FUNC(XTEMP_VECTOR, YTEMP_VECTOR, NROWS - (k + 1), 0)(6, 1)
            'If j = 1 Then: Must be more than -2.86 (5%) or -3.43 (1%) to proceed
            'with the regression; else the original series is I(0), so abort.
            'If j > 1 Then: Needs to be less than -2.86 (5%) or more than -3.43
            '(1%) to proceed with the regression. If so, then we do not reject the
            'hypothesis that the difference series is I(0), so conclude the original
            'series is I(1)
        Next j
    Next k

'-------------------------------------------------------------------------------
Else
'-------------------------------------------------------------------------------
    YDATA_VECTOR = YDATA_RNG
    If UBound(YDATA_VECTOR, 1) = 1 Then
        YDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(YDATA_VECTOR)
    End If
    RESIDUAL_MATRIX = RESIDUALS_REGRESSION_FUNC(XDATA_MATRIX, YDATA_VECTOR, True, 0)

    ReDim TEMP_MATRIX(0 To nLAGS, 1 To 2)
    TEMP_MATRIX(0, 1) = "LAGS"
    TEMP_MATRIX(0, 2) = "RATIO"

    For k = 1 To nLAGS
        TEMP_MATRIX(k, 1) = k
        ReDim YTEMP_VECTOR(1 To NROWS - k, 1 To 1)
        For j = 1 To NCOLUMNS
            XTEMP_VECTOR = MATRIX_GET_COLUMN_FUNC(RESIDUAL_MATRIX, j, 1)
            For i = 1 To NROWS - k
                YTEMP_VECTOR(i, 1) = XTEMP_VECTOR(i + 1, 1) - XTEMP_VECTOR(i, 1)
            Next i
            TEMP_MATRIX(k, 2) = COINTEGRATION_EG_STATISTICS_FUNC(XTEMP_VECTOR, YTEMP_VECTOR, NROWS - k, 0)(6, 1)
            'Needs to be less than -2.86 (at the 5% level) or -3.43 (at the 1% level) to
            'conclude that the residuals are I(0), so the cointegration equation is legitimate.
        Next j
    Next k
'-------------------------------------------------------------------------------
End If
'-------------------------------------------------------------------------------

COINTEGRATION_EG_TEST_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
COINTEGRATION_EG_TEST_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : COINTEGRATION_EG_STATISTICS_FUNC
'DESCRIPTION   : Calculating cointegration statistics for time series. The
'routine tests if the input series are I(1), and if the cointegration regression
'residuals are I(0), according to the Engle-Granger methodology.
'The routine outputs the statistics and comments on each statistic to guide
'the decision as to the cointegration relationship.

'LIBRARY       : STATISTICS
'GROUP         : COINTEGRATION
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function COINTEGRATION_EG_STATISTICS_FUNC(ByRef XDATA_RNG As Variant, _
ByRef YDATA_RNG As Variant, _
Optional ByVal NROWS As Long = 0, _
Optional ByVal OUTPUT As Integer = 1)

Dim i As Long
Dim j As Long
Dim k As Long

Dim NCOLUMNS As Long

Dim EY_VAL As Double
Dim EYY_VAL As Double
Dim YSIGMA_VAL As Double

Dim EX_VECTOR As Variant
Dim EXX_MATRIX As Variant
Dim EXY_VECTOR As Variant

Dim XDATA_MATRIX As Variant
Dim YDATA_VECTOR As Variant

Dim TEMP_VECTOR As Variant

On Error GoTo ERROR_LABEL

XDATA_MATRIX = XDATA_RNG
YDATA_VECTOR = YDATA_RNG
If UBound(YDATA_VECTOR, 1) = 1 Then
    YDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(YDATA_VECTOR)
End If
If NROWS = 0 Then: NROWS = UBound(YDATA_VECTOR, 1)

NCOLUMNS = UBound(XDATA_MATRIX, 2)

ReDim EX_VECTOR(1 To 1, 1 To NCOLUMNS)
ReDim EXX_MATRIX(1 To NCOLUMNS, 1 To NCOLUMNS)
ReDim EXY_VECTOR(1 To 1, 1 To NCOLUMNS)

For i = 1 To NROWS
    For j = 1 To NCOLUMNS
        EX_VECTOR(1, j) = EX_VECTOR(1, j) + XDATA_MATRIX(i, j)
        EXY_VECTOR(1, j) = EXY_VECTOR(1, j) + XDATA_MATRIX(i, j) * YDATA_VECTOR(i, 1)
        For k = 1 To NCOLUMNS
            EXX_MATRIX(j, k) = EXX_MATRIX(j, k) + XDATA_MATRIX(i, j) * XDATA_MATRIX(i, k)
        Next k
    Next j
    EY_VAL = EY_VAL + YDATA_VECTOR(i, 1)
    EYY_VAL = EYY_VAL + YDATA_VECTOR(i, 1) ^ 2
Next i

For j = 1 To NCOLUMNS
    EX_VECTOR(1, j) = EX_VECTOR(1, j) / NROWS
    EXY_VECTOR(1, j) = EXY_VECTOR(1, j) / NROWS
    For k = 1 To NCOLUMNS
        EXX_MATRIX(j, k) = EXX_MATRIX(j, k) / NROWS
    Next k
Next j

EY_VAL = EY_VAL / NROWS
EYY_VAL = EYY_VAL / NROWS

YSIGMA_VAL = Sqr(EYY_VAL - EY_VAL ^ 2)

'-----------------------------------------------------------------------------------
Select Case OUTPUT
'-----------------------------------------------------------------------------------
Case 0
'-----------------------------------------------------------------------------------
    ReDim TEMP_VECTOR(1 To 6, 1 To NCOLUMNS)
    For j = 1 To NCOLUMNS
        TEMP_VECTOR(1, j) = YSIGMA_VAL
        TEMP_VECTOR(2, j) = Sqr(EXX_MATRIX(j, j) - EX_VECTOR(1, j) ^ 2) 'SIGMA_X
        TEMP_VECTOR(3, j) = EXY_VECTOR(1, j) - EX_VECTOR(1, j) * EY_VAL 'SIGMA_XY
        TEMP_VECTOR(4, j) = TEMP_VECTOR(3, j) / TEMP_VECTOR(2, j) ^ 2 'BETA
        TEMP_VECTOR(5, j) = Sqr(1 / (NROWS - 2) * (YSIGMA_VAL ^ 2 / TEMP_VECTOR(2, j) ^ 2 - (TEMP_VECTOR(3, j) / TEMP_VECTOR(2, j) ^ 2) ^ 2)) 'SE BETA
        TEMP_VECTOR(6, j) = TEMP_VECTOR(4, j) / TEMP_VECTOR(5, j) 'RATIO
    Next j
    COINTEGRATION_EG_STATISTICS_FUNC = TEMP_VECTOR
'-----------------------------------------------------------------------------------
Case Else 'multi-factor E-G methodology
'-----------------------------------------------------------------------------------
    NCOLUMNS = UBound(EXX_MATRIX, 2)
    ReDim XDATA_MATRIX(1 To NCOLUMNS + 1, 1 To NCOLUMNS + 1)
    ReDim YDATA_VECTOR(1 To NCOLUMNS + 1, 1 To 1)
    For j = 1 To NCOLUMNS
        For k = 1 To NCOLUMNS
            XDATA_MATRIX(j, k) = EXX_MATRIX(j, k)
        Next k
        XDATA_MATRIX(NCOLUMNS + 1, j) = EX_VECTOR(1, j)
        XDATA_MATRIX(j, NCOLUMNS + 1) = EX_VECTOR(1, j)
        YDATA_VECTOR(j, 1) = EXY_VECTOR(1, j)
    Next j
    XDATA_MATRIX(NCOLUMNS + 1, NCOLUMNS + 1) = 1
    YDATA_VECTOR(NCOLUMNS + 1, 1) = EY_VAL
    XDATA_MATRIX = MATRIX_LU_LINEAR_SYSTEM_FUNC(XDATA_MATRIX, YDATA_VECTOR)
    NROWS = UBound(XDATA_MATRIX, 1)
    ReDim TEMP_VECTOR(1 To NROWS, 1 To 1)
    TEMP_VECTOR(1, 1) = XDATA_MATRIX(NROWS, 1)
    For i = 2 To NROWS
        TEMP_VECTOR(i, 1) = XDATA_MATRIX(i - 1, 1)
    Next i
    COINTEGRATION_EG_STATISTICS_FUNC = TEMP_VECTOR
'-----------------------------------------------------------------------------------
End Select
'-----------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
COINTEGRATION_EG_STATISTICS_FUNC = Err.number
End Function
