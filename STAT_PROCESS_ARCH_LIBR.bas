Attribute VB_Name = "STAT_PROCESS_ARCH_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : ARCH_ONE_DRAW_FUNC
'DESCRIPTION   : One Draw Experiments for Tests of First-Order Autocorrelation
'With Lagged Dependent Variable

'LIBRARY       : STATISTICS
'GROUP         : AR
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function ARCH_ONE_DRAW_FUNC(ByVal B0_VAL As Double, _
ByVal B1_VAL As Double, _
ByVal RHO_VAL As Double, _
ByVal SIGMA_ERROR_VAL As Double, _
ByVal MIN_TENOR As Double, _
ByVal MAX_TENOR As Double, _
Optional ByVal DELTA_TENOR As Double = 1, _
Optional ByVal OUTPUT As Integer = 0)

Dim i As Long
Dim j As Long
Dim NROWS As Long

Dim T_VAL As Double
Dim OLS_MATRIX As Variant
Dim TEMP_MATRIX As Variant
Dim TEMP_VECTOR As Variant
Dim RESID_VECTOR As Variant

Dim XTEMP_VECTOR As Variant
Dim YTEMP_VECTOR As Variant

'Yt = Bo + B1 x Yt-1 + Et
'Et = pEt-1 + Vt

'Bo = B0_VAL
'B1 = B1_VAL
'p = RHO_VAL
'V = SIGMA_ERROR_VAL
't = TENOR (periods in years, months, weeks or days)

'STEPS INVOLVED:
' (1) CREATE DATA MATRIX
' (2) RUN OLS ON DATA MATRIX
' (3) COMPUTE RESIDUALS
' (4) COMPUTE RESIDUALS REGRESSION ON X'S LAGGED RESIDS TEST STAT
' (5) COMPUTE TEST STAT, P VALUE
' (6) COMPUTE DW P-VALUE

On Error GoTo ERROR_LABEL

NROWS = (MAX_TENOR - MIN_TENOR) / DELTA_TENOR + 1
ReDim TEMP_MATRIX(0 To NROWS, 1 To 10)

ReDim XTEMP_VECTOR(1 To NROWS, 1 To 1)
ReDim YTEMP_VECTOR(1 To NROWS, 1 To 1)

ReDim TEMP_VECTOR(1 To NROWS - 1, 1 To 2)
ReDim RESID_VECTOR(1 To NROWS - 1, 1 To 1)

TEMP_MATRIX(0, 1) = "PERIOD"
TEMP_MATRIX(0, 2) = "YT-1"
TEMP_MATRIX(0, 3) = "VT"
TEMP_MATRIX(0, 4) = "ET"
TEMP_MATRIX(0, 5) = "YT"
TEMP_MATRIX(0, 6) = "EST. YT"
TEMP_MATRIX(0, 7) = "RESIDUAL"
TEMP_MATRIX(0, 8) = "XT"
TEMP_MATRIX(0, 9) = "RESIDUAL T-1"
TEMP_MATRIX(0, 10) = "(RESIDUAL [T] - RESIDUAL [T-1])^2"

T_VAL = MIN_TENOR
TEMP_MATRIX(1, 1) = T_VAL
TEMP_MATRIX(1, 2) = 0
TEMP_MATRIX(1, 3) = 0
TEMP_MATRIX(1, 4) = RANDOM_NORMAL_FUNC(0, SIGMA_ERROR_VAL * Sqr((1 - RHO_VAL ^ 2)), 0)
TEMP_MATRIX(1, 5) = 0

T_VAL = T_VAL + DELTA_TENOR

j = 1
For i = 2 To NROWS
    'We deal with the initial conditions problem by starting the process in the past …
    TEMP_MATRIX(i, 1) = T_VAL
    TEMP_MATRIX(i, 2) = TEMP_MATRIX(i - 1, 5)
    TEMP_MATRIX(i, 3) = RANDOM_NORMAL_FUNC(0, SIGMA_ERROR_VAL, 0)
    TEMP_MATRIX(i, 4) = RHO_VAL * TEMP_MATRIX(i - 1, 4) + TEMP_MATRIX(i, 3)
    TEMP_MATRIX(i, 5) = B0_VAL + B1_VAL * TEMP_MATRIX(i, 2) + TEMP_MATRIX(i, 4)
    If T_VAL > 0 Then
        XTEMP_VECTOR(j, 1) = TEMP_MATRIX(i, 2)
        YTEMP_VECTOR(j, 1) = TEMP_MATRIX(i, 5)
        j = j + 1
    End If
    T_VAL = T_VAL + DELTA_TENOR
Next i

OLS_MATRIX = REGRESSION_SIMPLE_COEF_FUNC(XTEMP_VECTOR, YTEMP_VECTOR)
'Compute residuals regression on X's lagged resids test stat
B1_VAL = OLS_MATRIX(1, 1)
B0_VAL = OLS_MATRIX(2, 1)

TEMP_MATRIX(1, 6) = ""
TEMP_MATRIX(1, 7) = ""
TEMP_MATRIX(1, 8) = ""
TEMP_MATRIX(1, 9) = ""
TEMP_MATRIX(1, 10) = ""

T_VAL = MIN_TENOR

j = 1
For i = 1 To NROWS
'---------------------------------------------------------------------------------------
    If (T_VAL > 0) Then
'---------------------------------------------------------------------------------------
        If (T_VAL > DELTA_TENOR) Then
            TEMP_MATRIX(i, 6) = B0_VAL + B1_VAL * TEMP_MATRIX(i, 2)
            TEMP_MATRIX(i, 7) = TEMP_MATRIX(i, 5) - TEMP_MATRIX(i, 6)
            TEMP_MATRIX(i, 8) = TEMP_MATRIX(i, 2)
            TEMP_MATRIX(i, 9) = TEMP_MATRIX(i - 1, 7)
            TEMP_MATRIX(i, 10) = (TEMP_MATRIX(i, 7) - TEMP_MATRIX(i, 9)) ^ 2
        Else
            TEMP_MATRIX(i, 6) = B0_VAL + B1_VAL * TEMP_MATRIX(i, 2)
            TEMP_MATRIX(i, 7) = TEMP_MATRIX(i, 5) - TEMP_MATRIX(i, 6)
            TEMP_MATRIX(i, 8) = TEMP_MATRIX(i, 2)
            TEMP_MATRIX(i, 9) = 0
            TEMP_MATRIX(i, 10) = 0
        End If
        If (T_VAL > 1) Then
            RESID_VECTOR(j, 1) = TEMP_MATRIX(i, 7)
            TEMP_VECTOR(j, 1) = TEMP_MATRIX(i, 8) 'Yt-1
            TEMP_VECTOR(j, 2) = TEMP_MATRIX(i, 9) 'Resid t-1
            j = j + 1
        End If
'---------------------------------------------------------------------------------------
    Else
'---------------------------------------------------------------------------------------
        TEMP_MATRIX(i, 6) = 0
        TEMP_MATRIX(i, 7) = 0
        TEMP_MATRIX(i, 8) = 0
        TEMP_MATRIX(i, 9) = 0
        TEMP_MATRIX(i, 10) = 0
'---------------------------------------------------------------------------------------
    End If
'---------------------------------------------------------------------------------------
    T_VAL = T_VAL + DELTA_TENOR
Next i

Select Case OUTPUT
Case 0 'AR Model
    ARCH_ONE_DRAW_FUNC = TEMP_MATRIX
'Case 1 'Autocorrelation - Test
 '   ARCH_ONE_DRAW_FUNC = AUTO_CORREL_LAG_TEST_FUNC(TEMP_VECTOR, RESID_VECTOR, True)
Case 1 'Residual Test for AR(1)
    ARCH_ONE_DRAW_FUNC = REGRESSION_LS1_FUNC(TEMP_VECTOR, RESID_VECTOR, True, 0, 0)
Case Else
    ARCH_ONE_DRAW_FUNC = REGRESSION_LS1_FUNC(XTEMP_VECTOR, YTEMP_VECTOR, True, 0, 0)
End Select

Exit Function
ERROR_LABEL:
ARCH_ONE_DRAW_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : ARMA_SIMULATION_FUNC
'DESCRIPTION   : ARMA(p,q) auto correlation simulation
'LIBRARY       : STATISTICS
'GROUP         : AR
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function ARMA_SIMULATION_FUNC(ByVal NSIZE As Long, _
ByVal LOG_MEAN_VAL As Double, _
ByVal LOG_SIGMA_VAL As Double, _
ByVal AR1_LAG As Double, _
ByVal AR2_LAG As Double, _
ByVal MA1_LAG As Double, _
ByVal MA2_LAG As Double, _
Optional ByVal nLOOPS As Long = 100)

Dim i As Long
Dim j As Long

Dim TEMP_FACTOR As Double
Dim TEMP_MATRIX As Variant
Dim RANDOM_MATRIX As Variant

On Error GoTo ERROR_LABEL

ReDim TEMP_MATRIX(1 To NSIZE, 1 To nLOOPS)

TEMP_FACTOR = LOG_MEAN_VAL * (1 - AR1_LAG - AR2_LAG)
RANDOM_MATRIX = MATRIX_RANDOM_NORMAL_FUNC(NSIZE, nLOOPS, 0, 0, LOG_SIGMA_VAL, 0)

For j = 1 To nLOOPS
    TEMP_MATRIX(1, j) = TEMP_FACTOR + RANDOM_MATRIX(1, j)
    TEMP_MATRIX(2, j) = TEMP_FACTOR + RANDOM_MATRIX(2, j) + AR1_LAG * TEMP_MATRIX(1, j) + MA1_LAG * RANDOM_MATRIX(1, j)
    For i = 3 To NSIZE
        TEMP_MATRIX(i, j) = TEMP_FACTOR + RANDOM_MATRIX(i, j) + AR1_LAG * TEMP_MATRIX(i - 1, j) + AR2_LAG * TEMP_MATRIX(i - 2, j) + MA1_LAG * RANDOM_MATRIX(i - 1, j) + MA2_LAG * RANDOM_MATRIX(i - 2, j)
    Next i
Next j

ARMA_SIMULATION_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
ARMA_SIMULATION_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : AR_DRIFT_FUNC
'DESCRIPTION   : This functions draw Future Values of a set for Trend
'AR(1) and RW w/ Drift Models
'LIBRARY       : STATISTICS
'GROUP         : AR
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************


Function AR_DRIFT_FUNC(ByVal B0_VAL As Double, _
ByVal B1_VAL As Double, _
ByVal SIGMA_ERROR0_VAL As Double, _
ByVal RHO_VAL As Double, _
Optional ByVal INITIAL_ERROR0_VAL As Double = -5.80820911223352E-02, _
Optional ByVal SIGMA_ERROR1_VAL As Double = 2.32570160061725E-02, _
Optional ByVal INITIAL_ERROR1_VAL As Double = 9.24936875894366, _
Optional ByVal BETA_ERROR1_VAL As Double = 3.37084949714747E-02, _
Optional ByVal MIN_TENOR As Double = 0, _
Optional ByVal MAX_TENOR As Double = 30, _
Optional ByVal DELTA_TENOR As Double = 1)

'------------------------------------------------------------------------------------------
'INITIAL_ERROR0_VAL: This is a parameter that is assumed known with certainty
'--------------------------------INITIAL_ERROR1_VAL: epsilon---------------------------------------
'|------------------Autocorrelated Trend and Random Walk with Drift------------------------|
'                    Yt = Bo + B1t + et, t = 1,....T ; AR1 Trend Y
'                    Yt = y1 + Yt-1 + vt ; Random Walk Y
'                    Et = pEt-1 + Vt
'|-----------------------------------------------------------------------------------------|

Dim i As Long
Dim NROWS As Long

Dim T_VAL As Double
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

NROWS = (MAX_TENOR - MIN_TENOR) / DELTA_TENOR + 1
ReDim TEMP_MATRIX(0 To NROWS, 1 To 9)

TEMP_MATRIX(0, 1) = "TENOR"
    
'-----------------------Expected Values----------------------
TEMP_MATRIX(0, 2) = "EV: AR1 E[et]"
TEMP_MATRIX(0, 3) = "EV: AR1 E[Yt]"
TEMP_MATRIX(0, 4) = "EV: RW E[Yt]"
'-----------------------Realized Values----------------------
TEMP_MATRIX(0, 5) = "RV: et"
TEMP_MATRIX(0, 6) = "RV: AR(1) Yt"
TEMP_MATRIX(0, 7) = "RV: RW Yt"
'-----------------------Forecast Errors----------------------
TEMP_MATRIX(0, 8) = "FE: AR(1)"
TEMP_MATRIX(0, 9) = "FE: RW"
'------------------------------------------------------------
For i = 1 To 9: TEMP_MATRIX(1, i) = "": Next i
TEMP_MATRIX(1, 1) = 0
TEMP_MATRIX(1, 5) = INITIAL_ERROR0_VAL 'This is a parameter that is assumed known with certainty
T_VAL = MIN_TENOR + DELTA_TENOR
For i = 2 To NROWS
'----------------------------------------------------------------------------------------
    TEMP_MATRIX(i, 1) = T_VAL
    TEMP_MATRIX(i, 2) = TEMP_MATRIX(1, 5) * RHO_VAL ^ TEMP_MATRIX(i, 1)
    TEMP_MATRIX(i, 3) = B0_VAL + B1_VAL * TEMP_MATRIX(i, 1) + TEMP_MATRIX(i, 2)
    TEMP_MATRIX(i, 4) = INITIAL_ERROR1_VAL + BETA_ERROR1_VAL * TEMP_MATRIX(i, 1)
    TEMP_MATRIX(i, 5) = RHO_VAL * TEMP_MATRIX(i - 1, 5) + RANDOM_NORMAL_FUNC(0, SIGMA_ERROR0_VAL, 0)
'--------------------Autocorrelated Trend and Random Walk with Drift---------------------
    TEMP_MATRIX(i, 6) = B0_VAL + B1_VAL * TEMP_MATRIX(i, 1) + TEMP_MATRIX(i, 5)
    If i <> 2 Then
        TEMP_MATRIX(i, 7) = BETA_ERROR1_VAL + TEMP_MATRIX(i - 1, 7) + RANDOM_NORMAL_FUNC(0, SIGMA_ERROR1_VAL, 0)
    Else
        TEMP_MATRIX(i, 7) = INITIAL_ERROR1_VAL + BETA_ERROR1_VAL + RANDOM_NORMAL_FUNC(0, SIGMA_ERROR1_VAL, 0)
    End If
'-----------Forecast Errors for Autocorrelated Trend and Random Walk with Drift----------
    TEMP_MATRIX(i, 8) = TEMP_MATRIX(i, 6) - TEMP_MATRIX(i, 3)
    TEMP_MATRIX(i, 9) = TEMP_MATRIX(i, 7) - TEMP_MATRIX(i, 4)
'----------------------------------------------------------------------------------------
    T_VAL = T_VAL + DELTA_TENOR
'----------------------------------------------------------------------------------------
Next i

AR_DRIFT_FUNC = TEMP_MATRIX
  
Exit Function
ERROR_LABEL:
AR_DRIFT_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : AR_PARTIAL_LOG_TREND_FUNC
'DESCRIPTION   : Partial Log Regression Model
'This routine contains a linear and logarithmic trend models.
'The function produces Predicted data series, and demonstrates two correction methods
'for converting Predicted Ln Data to the Predicted Level of Data.
'LIBRARY       : STATISTICS
'GROUP         : REGRESSION_LOG
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'REFERENCE:
'Diebold, Francis X. (2001) Elements of Forecasting, 2nd Edition, pp.329-331
'************************************************************************************
'************************************************************************************

Function AR_PARTIAL_LOG_TREND_FUNC(ByVal XDATA_RNG As Variant, _
ByVal YDATA_RNG As Variant, _
Optional ByVal OUTPUT As Integer = 0)

Dim i As Long
Dim NROWS As Long

Dim SLOPE As Double
Dim B0_VAL As Double

Dim CORREC_FACTOR As Double 'Correction Factor
Dim TREND_GROWTH As Double 'Ln Trend Growth Rate

Dim XDATA_VECTOR As Variant
Dim YDATA_VECTOR As Variant

Dim TEMP_MATRIX As Variant

Dim LINEAR_MATRIX As Variant
Dim LOG_MATRIX As Variant
Dim CORREC_MATRIX As Variant 'General Method of Correction
'in Going From Log to Levels

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

If UBound(YDATA_VECTOR, 1) <> UBound(XDATA_VECTOR, 1) Then: GoTo ERROR_LABEL

ReDim TEMP_MATRIX(0 To NROWS, 1 To 9)
TEMP_MATRIX(0, 1) = "INDEX"
TEMP_MATRIX(0, 2) = "DATE"
TEMP_MATRIX(0, 3) = "DATA"
TEMP_MATRIX(0, 4) = "Ln(DATA)"
TEMP_MATRIX(0, 5) = "LINEAR PREDICTED"
TEMP_MATRIX(0, 6) = "LOG LINEAR PREDICTED"
TEMP_MATRIX(0, 7) = "EXP(PREDICTED)"
TEMP_MATRIX(0, 8) = "PREDICTED LN YDATA"
TEMP_MATRIX(0, 9) = "LOG LINEAR PREDICTED NORM CORREC"
For i = 1 To NROWS
    TEMP_MATRIX(i, 1) = i 'Index
    TEMP_MATRIX(i, 2) = XDATA_VECTOR(i, 1) 'Date
    TEMP_MATRIX(i, 3) = YDATA_VECTOR(i, 1) ' Data
    TEMP_MATRIX(i, 4) = Log(TEMP_MATRIX(i, 3)) ' Log-Data
    XDATA_VECTOR(i, 1) = TEMP_MATRIX(i, 1)
Next i

LINEAR_MATRIX = REGRESSION_LS1_FUNC(XDATA_VECTOR, YDATA_VECTOR, True, 0, 0)  'Linear OLS

B0_VAL = LINEAR_MATRIX(6, 2)
SLOPE = LINEAR_MATRIX(7, 2)

For i = 1 To NROWS
    TEMP_MATRIX(i, 5) = B0_VAL + SLOPE * TEMP_MATRIX(i, 1) 'Linear Predicted
    YDATA_VECTOR(i, 1) = TEMP_MATRIX(i, 4)
Next i
LOG_MATRIX = REGRESSION_LS1_FUNC(XDATA_VECTOR, YDATA_VECTOR, True, 0, 0) 'Logarithmic OLS

B0_VAL = LOG_MATRIX(6, 2): SLOPE = LOG_MATRIX(7, 2)
TREND_GROWTH = Exp(SLOPE) - 1
CORREC_FACTOR = Exp(LOG_MATRIX(3, 2) ^ 2 / 2)

For i = 1 To NROWS
    TEMP_MATRIX(i, 8) = B0_VAL + SLOPE * TEMP_MATRIX(i, 1) 'Predicted Ln
    TEMP_MATRIX(i, 7) = Exp(TEMP_MATRIX(i, 8)) 'Exp (Predicted Ln)
    
    XDATA_VECTOR(i, 1) = TEMP_MATRIX(i, 7)
    YDATA_VECTOR(i, 1) = TEMP_MATRIX(i, 3)
Next i

'---------------Normal Distribution Method for Correction Factor--------------------
CORREC_MATRIX = REGRESSION_LS1_FUNC(XDATA_VECTOR, YDATA_VECTOR, False, 0, 0) 'Logarithmic OLS
'General Method of Correction in Going From Log to Levels
'-----------------------------------------------------------------------------------

For i = 1 To NROWS
    TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 7) * CORREC_MATRIX(6, 2)
    TEMP_MATRIX(i, 9) = Exp(TEMP_MATRIX(i, 8)) * CORREC_FACTOR
Next i

Select Case OUTPUT
Case 0
    AR_PARTIAL_LOG_TREND_FUNC = TEMP_MATRIX
Case 1
    AR_PARTIAL_LOG_TREND_FUNC = TREND_GROWTH
Case 2
    AR_PARTIAL_LOG_TREND_FUNC = CORREC_FACTOR
Case 3
    AR_PARTIAL_LOG_TREND_FUNC = LINEAR_MATRIX
Case 4
    AR_PARTIAL_LOG_TREND_FUNC = LOG_MATRIX
Case 5
    AR_PARTIAL_LOG_TREND_FUNC = CORREC_MATRIX
Case Else
    AR_PARTIAL_LOG_TREND_FUNC = Array(TEMP_MATRIX, TREND_GROWTH, CORREC_FACTOR, LINEAR_MATRIX, LOG_MATRIX, CORREC_MATRIX)
End Select

Exit Function
ERROR_LABEL:
AR_PARTIAL_LOG_TREND_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : AR_FULL_LOG_TREND_FUNC
'DESCRIPTION   : Full Log Regression Model

'This routine produces Predicted data series, and demonstrates two correction methods for
'converting Predicted Ln Data to the Predicted Level of Data. The RandomWalkvsTrend fits a
'random walk (with drift) model to Real GDP and compares the random walk to an AR(1)
'model which includes a trend.

'LIBRARY       : STATISTICS
'GROUP         : REGRESSION_LOG
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'REFERENCE:
'Diebold, Francis X. (2001) Elements of Forecasting, 2nd Edition, pp.329-331
'************************************************************************************
'************************************************************************************

Function AR_FULL_LOG_TREND_FUNC(ByVal XDATA_RNG As Variant, _
ByVal YDATA_RNG As Variant, _
Optional ByVal FORWARD_PERIODS As Long = 6, _
Optional ByVal INIT_INTER As Double = 1, _
Optional ByVal OUTPUT As Integer = 0)

Dim i As Long
Dim j As Long
Dim NROWS As Long

Dim SLOPE As Double
Dim B0_VAL As Double

Dim MEAN_VAL As Double
Dim SIGMA_VAL As Double
Dim TEMP_SUM As Double

Dim TREND_GROWTH As Double 'Transformed Ln Trend Growth Rate
Dim DRIFT_GROWTH As Double 'Random Walk Drift Growth Rate

Dim XDATA_VECTOR As Variant
Dim YDATA_VECTOR As Variant

Dim TEMP_MATRIX As Variant

Dim TREND_MATRIX As Variant 'LN_TREND_OLS
Dim RHO_VAL_MATRIX As Variant 'RHO_VAL_OLS
Dim RANDOM_MATRIX As Variant 'RANDOM_OLS
Dim LAG_MATRIX As Variant 'LAG_OLS
Dim RESIDUAL_MATRIX As Variant 'RESIDUAL_OLS

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

If UBound(YDATA_VECTOR, 1) <> UBound(XDATA_VECTOR, 1) Then: GoTo ERROR_LABEL

'----------------------------------------------------------------------
If FORWARD_PERIODS > NROWS Then
    ReDim TEMP_MATRIX(0 To FORWARD_PERIODS, 1 To 27)
Else
    ReDim TEMP_MATRIX(0 To NROWS, 1 To 27)
End If
For i = 1 To UBound(TEMP_MATRIX, 1): For j = 1 To UBound(TEMP_MATRIX, 2): _
TEMP_MATRIX(i, j) = "": Next j: Next i
'----------------------------------------------------------------------
TEMP_MATRIX(0, 1) = "INDEX"
TEMP_MATRIX(0, 2) = "DATE"
TEMP_MATRIX(0, 3) = "DATA"
TEMP_MATRIX(0, 4) = "Ln(DATA)"
TEMP_MATRIX(0, 5) = "PREDICTED Ln(DATA)"
TEMP_MATRIX(0, 6) = "LOG RESIDUAL"
TEMP_MATRIX(0, 7) = "LAG LOG RESIDUAL"
TEMP_MATRIX(0, 8) = "DLn(DATA)"
TEMP_MATRIX(0, 9) = "B0_VAL"
TEMP_MATRIX(0, 10) = "RANDOM WALK RESID"
TEMP_MATRIX(0, 11) = "LAG RANDOM WALK RESID"
TEMP_MATRIX(0, 12) = "TRANSF Ln(DATA)"
TEMP_MATRIX(0, 13) = "TRANSF B0_VAL"
TEMP_MATRIX(0, 14) = "TRANSF INDEX"
TEMP_MATRIX(0, 15) = "TRANSF RESID"
'--------------------------Point Forecasts----------------------------
TEMP_MATRIX(0, 16) = "PF:INDEX"
TEMP_MATRIX(0, 17) = "PF:FOREC Ln TREND"
TEMP_MATRIX(0, 18) = "PF:Ln RESID"
TEMP_MATRIX(0, 19) = "PF:FOREC Ln RANDOM WALK"
'-----------------------SE Assuming Known Parameters------------------
TEMP_MATRIX(0, 20) = "SKP:RANDOM WALK SE FOREC"
TEMP_MATRIX(0, 21) = "SKP:Ln TREND SE FOREC"
TEMP_MATRIX(0, 22) = "SKP:Ln TREND VaR FOREC"
'------------------------SE due to Estimation Error-------------------
TEMP_MATRIX(0, 23) = "SEE:EST ERROR"
TEMP_MATRIX(0, 24) = "SEE:RANDOM WALK SE"
TEMP_MATRIX(0, 25) = "SEE:Ln TREND SE"
'---------------------Overall SE Forecasted Ln YDATA------------------
TEMP_MATRIX(0, 26) = "OSE:RANDOM WALK SE"
TEMP_MATRIX(0, 27) = "OSE:Ln TREND SE"
'-----------------------------------------------------------------------------

TEMP_SUM = 0
For i = 1 To NROWS
    TEMP_MATRIX(i, 1) = i 'Index
    TEMP_SUM = TEMP_SUM + TEMP_MATRIX(i, 1)
    TEMP_MATRIX(i, 2) = XDATA_VECTOR(i, 1) 'Date
    TEMP_MATRIX(i, 3) = YDATA_VECTOR(i, 1) ' Data
    TEMP_MATRIX(i, 4) = Log(TEMP_MATRIX(i, 3)) ' Log-Data
    
    XDATA_VECTOR(i, 1) = TEMP_MATRIX(i, 1)
    YDATA_VECTOR(i, 1) = TEMP_MATRIX(i, 4)
Next i
MEAN_VAL = TEMP_SUM / NROWS
TREND_MATRIX = REGRESSION_LS1_FUNC(XDATA_VECTOR, YDATA_VECTOR, True, 0, 0) 'Ln Trend

B0_VAL = TREND_MATRIX(6, 2)
SLOPE = TREND_MATRIX(7, 2)

TEMP_SUM = 0
For i = 1 To NROWS
    TEMP_SUM = TEMP_SUM + (MEAN_VAL - TEMP_MATRIX(i, 1)) ^ 2
    TEMP_MATRIX(i, 5) = B0_VAL + SLOPE * TEMP_MATRIX(i, 1) 'Linear Predicted
    TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 4) - TEMP_MATRIX(i, 5)
Next i
SIGMA_VAL = (TEMP_SUM / (NROWS - 1)) ^ 0.5

i = 1
For j = 7 To 9: TEMP_MATRIX(i, j) = "": Next j
ReDim XDATA_VECTOR(1 To NROWS - 1, 1 To 1)
ReDim YDATA_VECTOR(1 To NROWS - 1, 1 To 1)
For i = 2 To NROWS
    TEMP_MATRIX(i, 7) = TEMP_MATRIX(i - 1, 6)
    TEMP_MATRIX(i, 8) = TEMP_MATRIX(i, 4) - TEMP_MATRIX(i - 1, 4)
    TEMP_MATRIX(i, 9) = INIT_INTER
    
    XDATA_VECTOR(i - 1, 1) = TEMP_MATRIX(i, 7)
    YDATA_VECTOR(i - 1, 1) = TEMP_MATRIX(i, 6)
Next i
RHO_VAL_MATRIX = REGRESSION_LS1_FUNC(XDATA_VECTOR, YDATA_VECTOR, False, 0, 0) 'Estimated rho

For i = 2 To NROWS
    XDATA_VECTOR(i - 1, 1) = TEMP_MATRIX(i, 9)
    YDATA_VECTOR(i - 1, 1) = TEMP_MATRIX(i, 8)
Next i
RANDOM_MATRIX = REGRESSION_LS1_FUNC(XDATA_VECTOR, YDATA_VECTOR, False, 0, 0) 'Random Walk

i = 1
For j = 10 To 11: TEMP_MATRIX(i, j) = "": Next j
TEMP_MATRIX(i, 12) = TEMP_MATRIX(i, 4) / Sqr(1 - RHO_VAL_MATRIX(6, 2) ^ 2)
TEMP_MATRIX(i, 13) = 1 / Sqr(1 - RHO_VAL_MATRIX(6, 2) ^ 2)
TEMP_MATRIX(i, 14) = TEMP_MATRIX(i, 1) / Sqr(1 - RHO_VAL_MATRIX(6, 2) ^ 2)

i = 2
TEMP_MATRIX(i, 10) = TEMP_MATRIX(i, 8) - RANDOM_MATRIX(6, 2)
TEMP_MATRIX(i, 11) = ""
TEMP_MATRIX(i, 12) = TEMP_MATRIX(i, 4) - RHO_VAL_MATRIX(6, 2) * TEMP_MATRIX(i - 1, 4)
TEMP_MATRIX(i, 13) = 1 - RHO_VAL_MATRIX(6, 2)
TEMP_MATRIX(i, 14) = TEMP_MATRIX(i, 1) - RHO_VAL_MATRIX(6, 2) * TEMP_MATRIX(i - 1, 1)

ReDim XDATA_VECTOR(1 To NROWS - 2, 1 To 1)
ReDim YDATA_VECTOR(1 To NROWS - 2, 1 To 1)

For i = 3 To NROWS
    TEMP_MATRIX(i, 10) = TEMP_MATRIX(i, 8) - RANDOM_MATRIX(6, 2)
    TEMP_MATRIX(i, 11) = TEMP_MATRIX(i - 1, 10)
    TEMP_MATRIX(i, 12) = TEMP_MATRIX(i, 4) - RHO_VAL_MATRIX(6, 2) * TEMP_MATRIX(i - 1, 4)
    TEMP_MATRIX(i, 13) = 1 - RHO_VAL_MATRIX(6, 2)
    TEMP_MATRIX(i, 14) = TEMP_MATRIX(i, 1) - RHO_VAL_MATRIX(6, 2) * TEMP_MATRIX(i - 1, 1)

    XDATA_VECTOR(i - 2, 1) = TEMP_MATRIX(i, 11)
    YDATA_VECTOR(i - 2, 1) = TEMP_MATRIX(i, 10)
Next i

LAG_MATRIX = REGRESSION_LS1_FUNC(XDATA_VECTOR, YDATA_VECTOR, False, 0, 0)
'Resids on Lagged Resids

ReDim XDATA_VECTOR(1 To NROWS, 1 To 2)
ReDim YDATA_VECTOR(1 To NROWS, 1 To 1)

For i = 1 To NROWS
    XDATA_VECTOR(i, 1) = TEMP_MATRIX(i, 13)
    XDATA_VECTOR(i, 2) = TEMP_MATRIX(i, 14)
    YDATA_VECTOR(i, 1) = TEMP_MATRIX(i, 12)
Next i

RESIDUAL_MATRIX = REGRESSION_LS1_FUNC(XDATA_VECTOR, YDATA_VECTOR, False, 0, 0)
'Transformed Ln Trend

B0_VAL = RESIDUAL_MATRIX(6, 2)
SLOPE = RESIDUAL_MATRIX(7, 2)
For i = 1 To NROWS
    TEMP_MATRIX(i, 15) = TEMP_MATRIX(i, 4) - (B0_VAL + SLOPE * TEMP_MATRIX(i, 1))
Next i

TREND_GROWTH = Exp(SLOPE) - 1 'Transformed Ln Trend Growth Rate
DRIFT_GROWTH = Exp(RANDOM_MATRIX(6, 2)) - 1 'Random Walk Drift Growth Rate

i = 1
TEMP_MATRIX(i, 16) = TEMP_MATRIX(NROWS, 1)
TEMP_MATRIX(i, 17) = TEMP_MATRIX(NROWS, 4)
TEMP_MATRIX(i, 18) = TEMP_MATRIX(NROWS, 15)
TEMP_MATRIX(i, 19) = TEMP_MATRIX(i, 17)
For j = 20 To 27: TEMP_MATRIX(i, j) = "": Next j

For i = 2 To FORWARD_PERIODS
    TEMP_MATRIX(i, 16) = TEMP_MATRIX(i - 1, 16) + 1
    TEMP_MATRIX(i, 18) = RHO_VAL_MATRIX(6, 2) * TEMP_MATRIX(i - 1, 18)
    
    TEMP_MATRIX(i, 17) = B0_VAL + SLOPE * TEMP_MATRIX(i, 16) + TEMP_MATRIX(i, 18)
    
    TEMP_MATRIX(i, 19) = TEMP_MATRIX(i - 1, 19) + RANDOM_MATRIX(6, 2)
    TEMP_MATRIX(i, 20) = RANDOM_MATRIX(3, 2) * Sqr(TEMP_MATRIX(i, 16) - TEMP_MATRIX(1, 16))
    
    If i <> 2 Then
        TEMP_MATRIX(i, 22) = TEMP_MATRIX(i - 1, 22) + RHO_VAL_MATRIX(6, 2) ^ ((TEMP_MATRIX(i, 16) - 1 - TEMP_MATRIX(1, 16)) * 2) * RHO_VAL_MATRIX(3, 2) ^ 2
    Else
        TEMP_MATRIX(i, 22) = RHO_VAL_MATRIX(3, 2) ^ 2
    End If
    
    TEMP_MATRIX(i, 21) = Sqr(TEMP_MATRIX(i, 22))
    
    TEMP_MATRIX(i, 23) = Sqr((1 / NROWS) + (TEMP_MATRIX(i, 16) - MEAN_VAL) ^ 2 / (NROWS * SIGMA_VAL ^ 2))
    
    TEMP_MATRIX(i, 24) = RANDOM_MATRIX(3, 2) * TEMP_MATRIX(i, 23)
    TEMP_MATRIX(i, 25) = RESIDUAL_MATRIX(3, 2) * TEMP_MATRIX(i, 23)
    TEMP_MATRIX(i, 26) = Sqr(TEMP_MATRIX(i, 20) ^ 2 + TEMP_MATRIX(i, 24) ^ 2)
    TEMP_MATRIX(i, 27) = Sqr(TEMP_MATRIX(i, 25) ^ 2 + TEMP_MATRIX(i, 21) ^ 2)
Next i

'---------------------------------------------------------------------------------------
Select Case OUTPUT
'---------------------------------------------------------------------------------------
Case 0
'---------------------------------------------------------------------------------------
    AR_FULL_LOG_TREND_FUNC = TEMP_MATRIX
Case 1
    AR_FULL_LOG_TREND_FUNC = TREND_GROWTH 'Transformed Ln Trend Growth Rate
Case 2
    AR_FULL_LOG_TREND_FUNC = DRIFT_GROWTH 'Random Walk Drift Growth Rate
Case 3
    AR_FULL_LOG_TREND_FUNC = TREND_MATRIX 'Ln Trend
Case 4
    AR_FULL_LOG_TREND_FUNC = RHO_VAL_MATRIX 'Estimated Rho
Case 5
    AR_FULL_LOG_TREND_FUNC = RANDOM_MATRIX 'Random Walk
Case 6
    AR_FULL_LOG_TREND_FUNC = LAG_MATRIX 'Resids on Lagged Resids
Case 7
    AR_FULL_LOG_TREND_FUNC = RESIDUAL_MATRIX 'Transformed Ln Trend
Case Else
    AR_FULL_LOG_TREND_FUNC = Array(TEMP_MATRIX, TREND_GROWTH, DRIFT_GROWTH, TREND_MATRIX, RHO_VAL_MATRIX, RANDOM_MATRIX, LAG_MATRIX, RESIDUAL_MATRIX)
'---------------------------------------------------------------------------------------
End Select
'---------------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
AR_FULL_LOG_TREND_FUNC = Err.number
End Function
