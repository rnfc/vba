Attribute VB_Name = "STAT_PROCESS_JOHANSEN_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : JOHANSEN_TEST_FUNC

'DESCRIPTION   : 'Johansen test statistics used are for constant and no time trend


'LIBRARY       : STATISTICS
'GROUP         : JOHANSEN
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 14/02/2008
'************************************************************************************
'************************************************************************************

Function JOHANSEN_TEST_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal nLAGS As Long = 5)

'DATA_RNG --> Limit 12 Independent Variables
'nLAGS --> Specify no of lags for regression model

Dim i As Long
Dim j As Long

Dim ii As Long
Dim jj As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_SUM As Double
Dim LR_TRACE_VAL As Double
Dim LR_MAX_EIGEN_VAL As Double

Dim SKK_MATRIX As Variant
Dim SK0_MATRIX As Variant
Dim S00_MATRIX As Variant

Dim X_MATRIX As Variant
Dim X_DMEAN_MATRIX As Variant
Dim DX_DEMEAN_MATRIX As Variant
Dim DX_LAGGED_DEMEAN_MATRIX As Variant

Dim X_RESIDUAL_REGRESSION_MATRIX As Variant
Dim DX_RESIDUALS_REGRESSION_MATRIX As Variant

Dim TEMP_MATRIX As Variant
Dim EIGEN_VALUES_MATRIX As Variant
Dim CRITICAL_TRACE_MATRIX As Variant
Dim CRITICAL_EIGEN_MATRIX As Variant

On Error GoTo ERROR_LABEL

X_MATRIX = DATA_RNG
X_MATRIX = MATRIX_DEMEAN_FUNC(X_MATRIX)
TEMP_MATRIX = MATRIX_ELEMENTS_CONSECUTIVE_SUBTRACT_FUNC(X_MATRIX)
GoSub LAG_LINE
DX_LAGGED_DEMEAN_MATRIX = MATRIX_DEMEAN_FUNC(DX_LAGGED_DEMEAN_MATRIX)
DX_DEMEAN_MATRIX = MATRIX_DEMEAN_FUNC(MATRIX_GET_SUB_MATRIX_FUNC(TEMP_MATRIX, nLAGS + 1, -1, -1, -1))

TEMP_MATRIX = MMULT_FUNC(DX_LAGGED_DEMEAN_MATRIX, IMULT_FUNC(DX_LAGGED_DEMEAN_MATRIX, DX_DEMEAN_MATRIX, 0), 70)
'DX_FITTED_REGRESSION_MATRIX
    
DX_RESIDUALS_REGRESSION_MATRIX = MATRIX_ELEMENTS_SUBTRACT_FUNC(DX_DEMEAN_MATRIX, TEMP_MATRIX, 1, 1)
X_DMEAN_MATRIX = MATRIX_DEMEAN_FUNC(MATRIX_GET_SUB_MATRIX_FUNC(X_MATRIX, 2, UBound(X_MATRIX, 1) - nLAGS, -1, -1))

TEMP_MATRIX = MMULT_FUNC(DX_LAGGED_DEMEAN_MATRIX, IMULT_FUNC(DX_LAGGED_DEMEAN_MATRIX, X_DMEAN_MATRIX, 0), 70)
'X_FITTED_REGRESSION_MATRIX

X_RESIDUAL_REGRESSION_MATRIX = MATRIX_ELEMENTS_SUBTRACT_FUNC(X_DMEAN_MATRIX, TEMP_MATRIX, 1, 1)

SKK_MATRIX = MATRIX_ELEMENTS_DIVIDE_SCALAR_FUNC(MMULT_FUNC(MATRIX_TRANSPOSE_FUNC(X_RESIDUAL_REGRESSION_MATRIX), X_RESIDUAL_REGRESSION_MATRIX, 70), UBound(X_RESIDUAL_REGRESSION_MATRIX, 1))
SK0_MATRIX = MATRIX_ELEMENTS_DIVIDE_SCALAR_FUNC(MMULT_FUNC(MATRIX_TRANSPOSE_FUNC(X_RESIDUAL_REGRESSION_MATRIX), DX_RESIDUALS_REGRESSION_MATRIX, 70), UBound(X_RESIDUAL_REGRESSION_MATRIX, 1))

S00_MATRIX = MATRIX_ELEMENTS_DIVIDE_SCALAR_FUNC(MMULT_FUNC(MATRIX_TRANSPOSE_FUNC(DX_RESIDUALS_REGRESSION_MATRIX), DX_RESIDUALS_REGRESSION_MATRIX, 70), UBound(DX_RESIDUALS_REGRESSION_MATRIX, 1))

'SIGMA_MATRIX = MMULT_FUNC(MMULT_FUNC(SK0_MATRIX, MATRIX_INVERSE_FUNC( _
    S00_MATRIX, 2), 70), MATRIX_TRANSPOSE_FUNC(SK0_MATRIX), 70)

'ATEMP_MATRIX = MATRIX_INVERSE_FUNC(SKK_MATRIX,2)

'BTEMP_MATRIX = MMULT_FUNC(MATRIX_INVERSE_FUNC(SKK_MATRIX, 2), _
    MMULT_FUNC(MMULT_FUNC(SK0_MATRIX, MATRIX_INVERSE_FUNC(S00_MATRIX, 2), 70), _
    MATRIX_TRANSPOSE_FUNC(SK0_MATRIX), 70), 70)

'------------------------------------------------------------------------------------

TEMP_MATRIX = MMULT_FUNC(MATRIX_INVERSE_FUNC(SKK_MATRIX, 2), MMULT_FUNC(MMULT_FUNC(SK0_MATRIX, MATRIX_INVERSE_FUNC(S00_MATRIX, 2), 70), MATRIX_TRANSPOSE_FUNC(SK0_MATRIX), 70), 70)
EIGEN_VALUES_MATRIX = MATRIX_EIGEN_SQUARE_FUNC(TEMP_MATRIX, True, 0)
          
'------------------------------------------------------------------------------------
NROWS = UBound(X_RESIDUAL_REGRESSION_MATRIX, 1)
NCOLUMNS = UBound(X_RESIDUAL_REGRESSION_MATRIX, 2)

CRITICAL_TRACE_MATRIX = JOHANSEN_TEST_CRITICAL_VALUES_FUNC(3)
CRITICAL_EIGEN_MATRIX = JOHANSEN_TEST_CRITICAL_VALUES_FUNC(4)

ReDim TEMP_MATRIX(0 To UBound(X_MATRIX, 2), 1 To 10)
'johansen_results_trace
TEMP_MATRIX(0, 1) = "Trace H0: Rank<=x"
TEMP_MATRIX(0, 2) = "Trace Test statistic"
TEMP_MATRIX(0, 3) = "Trace Crit: 90%"
TEMP_MATRIX(0, 4) = "Trace Crit: 95%"
TEMP_MATRIX(0, 5) = "Trace Crit: 99%"
TEMP_MATRIX(0, 6) = "Eigen H0: Rank<=x"
TEMP_MATRIX(0, 7) = "Eigen Test statistic"
TEMP_MATRIX(0, 8) = "Eigen Crit: 90%"
TEMP_MATRIX(0, 9) = "Eigen Crit: 95%"
TEMP_MATRIX(0, 10) = "Eigen Crit: 99%"

'johansen_results_eigen
For i = 1 To UBound(X_MATRIX, 2)
    TEMP_SUM = 0
    For j = i To UBound(X_MATRIX, 2)
      TEMP_SUM = TEMP_SUM + Log(1 - EIGEN_VALUES_MATRIX(j, j))
    Next j
    LR_TRACE_VAL = -NROWS * TEMP_SUM
    LR_MAX_EIGEN_VAL = -NROWS * Log(1 - EIGEN_VALUES_MATRIX(i, i))
    
    TEMP_MATRIX(i, 1) = i - 1
    TEMP_MATRIX(i, 2) = LR_TRACE_VAL
    TEMP_MATRIX(i, 3) = CRITICAL_TRACE_MATRIX(NCOLUMNS - i + 1, 1)
    TEMP_MATRIX(i, 4) = CRITICAL_TRACE_MATRIX(NCOLUMNS - i + 1, 2)
    TEMP_MATRIX(i, 5) = CRITICAL_TRACE_MATRIX(NCOLUMNS - i + 1, 3)
    
    TEMP_MATRIX(i, 6) = i - 1
    TEMP_MATRIX(i, 7) = LR_MAX_EIGEN_VAL
    TEMP_MATRIX(i, 8) = CRITICAL_EIGEN_MATRIX(NCOLUMNS - i + 1, 1)
    TEMP_MATRIX(i, 9) = CRITICAL_EIGEN_MATRIX(NCOLUMNS - i + 1, 2)
    TEMP_MATRIX(i, 10) = CRITICAL_EIGEN_MATRIX(NCOLUMNS - i + 1, 3)
Next i

JOHANSEN_TEST_FUNC = TEMP_MATRIX

Exit Function
'------------------------------------------------------------------------------------
LAG_LINE:
'------------------------------------------------------------------------------------
    'Returns a matrix by laying out consecutive columns with the
    'lagged variables
    NROWS = UBound(TEMP_MATRIX, 1)
    NCOLUMNS = UBound(TEMP_MATRIX, 2)
    ReDim DX_LAGGED_DEMEAN_MATRIX(1 To NROWS - nLAGS, 1 To NCOLUMNS * nLAGS)
    ii = 1: jj = 1
    For j = 1 To NCOLUMNS * nLAGS
        For i = 1 To NROWS - nLAGS
            DX_LAGGED_DEMEAN_MATRIX(i, j) = TEMP_MATRIX(i + nLAGS - ii, jj)
        Next i
        ii = ii + 1
        If (ii > nLAGS) Then
           ii = 1
           jj = jj + 1
        End If
    Next j
'------------------------------------------------------------------------------------
Return
'------------------------------------------------------------------------------------
ERROR_LABEL:
JOHANSEN_TEST_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : JOHANSEN_TEST_CRITICAL_VALUES_FUNC
'DESCRIPTION   : Critical values for Johansens's test
'LIBRARY       : STATISTICS
'GROUP         : JOHANSEN
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 14/02/2008
'************************************************************************************
'************************************************************************************

Private Function JOHANSEN_TEST_CRITICAL_VALUES_FUNC( _
Optional ByVal OUTPUT As Integer = 0)

Dim i As Long
Dim ii As Long
Dim jj As Long

Dim DATA_STR As String
Dim TEMP_GROUP(1 To 6) As String
Dim TEMP_MATRIX(1 To 12, 1 To 3) As Variant
'C1 - 90% Confidence
'C2 - 95% Confidence
'C3 - 99% Confidence

On Error GoTo ERROR_LABEL

'-------------------------------------No deterministic part
TEMP_GROUP(1) = "2.9762,4.1296,6.9406|10.4741,12.3212,16.364|" & _
                "21.7781,24.2761,29.5147|37.0339,40.1749,46.5716|" & _
                "56.2839,60.0627,67.6367|79.5329,83.9383,92.7136|" & _
                "106.7351,111.7797,121.7375|137.9954,143.6691,154.7977|" & _
                "173.2292,179.5199,191.8122|212.4721,219.4051,232.8291|" & _
                "255.6732,263.2603,277.9962|302.9054,311.1288,326.9716|"
'Trace statistic
If OUTPUT = 1 Then: DATA_STR = TEMP_GROUP(1)
TEMP_GROUP(2) = "2.9762,4.1296,6.9406|9.4748,11.2246,15.0923|" & _
                "15.7175,17.7961,22.2519|21.837,24.1592,29.0609|" & _
                "27.916,30.4428,35.7359|33.9271,36.6301,42.2333|" & _
                "39.9085,42.7679,48.6606|45.893,48.8795,55.0335|" & _
                "51.8528,54.9629,61.3449|57.7954,61.0404,67.6415|" & _
                "63.7248,67.0756,73.8856|69.6513,73.0946,80.0937|"
'Maximum Eigen value statistic
If OUTPUT = 2 Then: DATA_STR = TEMP_GROUP(2)

'-----------------------------------------Constant term
TEMP_GROUP(3) = "2.7055,3.8415,6.6349|13.4294,15.4943,19.9349|" & _
                "27.0669,29.7961,35.4628|44.4929,47.8545,54.6815|" & _
                "65.8202,69.8189,77.8202|91.109,95.7542,104.9637|" & _
                "120.3673,125.6185,135.9825|153.6341,159.529,171.0905|" & _
                "190.8714,197.3772,210.0366|232.103,239.2468,253.2526|" & _
                "277.374,285.1402,300.2821|326.5354,334.9795,351.215|"
'Trace statistic
If OUTPUT = 3 Then: DATA_STR = TEMP_GROUP(3)

TEMP_GROUP(4) = "2.7055,3.8415,6.6349|12.2971,14.2639,18.52|" & _
                "18.8928,21.1314,25.865|25.1236,27.5858,32.7172|" & _
                "31.2379,33.8777,39.3693|37.2786,40.0763,45.8662|" & _
                "43.2947,46.2299,52.3069|49.2855,52.3622,58.6634|" & _
                "55.2412,58.4332,64.996|61.2041,64.504,71.2525|" & _
                "67.1307,70.5392,77.4877|73.0563,76.5734,83.7105|"
'Maximum Eigen value statistic
If OUTPUT = 4 Then: DATA_STR = TEMP_GROUP(4)

'-------------------------------------Constant and time trend
TEMP_GROUP(5) = "2.7055,3.8415,6.6349|16.1619,18.3985,23.1485|" & _
                "32.0645,35.0116,41.0815|51.6492,55.2459,62.5202|" & _
                "75.1027,79.3422,87.7748|102.4674,107.3429,116.9829|" & _
                "133.7852,139.278,150.0778|169.0618,175.1584,187.1891|" & _
                "208.3582,215.1268,228.2226|251.6293,259.0267,273.3838|" & _
                "298.8836,306.8988,322.4264|350.1125,358.719,375.3203|"
'Trace statistic
If OUTPUT = 5 Then: DATA_STR = TEMP_GROUP(5)

TEMP_GROUP(6) = "2.7055,3.8415,6.6349|15.0006,17.1481,21.7465|" & _
                "21.8731,24.2522,29.2631|28.2398,30.8151,36.193|" & _
                "34.4202,37.1646,42.8612|40.5244,43.4183,49.4095|" & _
                "46.5583,49.5875,55.8171|52.5858,55.7302,62.1741|" & _
                "58.5316,61.8051,68.503|64.5292,67.904,74.7434|" & _
                "70.463,73.9355,81.0678|76.4081,79.9878,87.2395|"
'Maximum Eigen value statistic
If OUTPUT > 5 Then: DATA_STR = TEMP_GROUP(6)

Select Case OUTPUT
Case 0
    JOHANSEN_TEST_CRITICAL_VALUES_FUNC = TEMP_GROUP
Case Else
    GoSub LOAD_MATRIX
    JOHANSEN_TEST_CRITICAL_VALUES_FUNC = TEMP_MATRIX
End Select

Exit Function
'--------------------------------------------------------------------------------
LOAD_MATRIX:
ii = 1
For i = 1 To 12
    jj = InStr(ii, DATA_STR, ",")
    TEMP_MATRIX(i, 1) = CDec(Mid(DATA_STR, ii, jj - ii))
    ii = jj + 1
    jj = InStr(ii, DATA_STR, ",")
    TEMP_MATRIX(i, 2) = CDec(Mid(DATA_STR, ii, jj - ii))
    ii = jj + 1
    jj = InStr(ii, DATA_STR, "|")
    TEMP_MATRIX(i, 3) = CDec(Mid(DATA_STR, ii, jj - ii))
    ii = jj + 1
Next i
Return
'--------------------------------------------------------------------------------
ERROR_LABEL:
JOHANSEN_TEST_CRITICAL_VALUES_FUNC = Err.number
End Function


