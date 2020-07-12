Attribute VB_Name = "STAT_PROCESS_ADF_LIBR"

'-------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------
Option Base 1
Option Explicit
'-------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------

'************************************************************************************
'************************************************************************************
'FUNCTION      : ADF_UNIT_ROOT_TEST_FUNC
'DESCRIPTION   :
'LIBRARY       : STATISTICS
'GROUP         :
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 05/26/2011
'************************************************************************************
'************************************************************************************

Function ADF_UNIT_ROOT_TEST_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal ROOT_MODE As Integer = 0, _
Optional ByVal TEST_MODE As Integer = 2, _
Optional ByVal LAG_AUTOMATIC_FLAG As Boolean = True, _
Optional ByVal LAG_LENGTH_MODE As Integer = 0, _
Optional ByRef MAX_LAGS As Long = 10, _
Optional ByVal NAME0_STR As String = "yt")

'ROOT_MODE
'0 - level
'1 - 1st Difference
'2 - 2nd Difference

'TEST_MODE
'0 - INTERCEPT_FLAG
'1 - TREND_VAL and INTERCEPT_FLAG
'2 - None

'LAG_LENGTH_MODE
'0: Akaike Info Criterion
'1: Schwartz Info Criterion
'2: Hannan-Quinn Criterion

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim NROWS As Long
Dim NOBS As Long
Dim NCOLUMNS As Long

Dim ROOT_INT As Long
Dim MODEL_INT As Long

Dim DF_VAL As Long
Dim TREND_VAL As Long
Dim EXOGEN1_VAL As Long
Dim EXOGEN2_VAL As Long

Dim MEAN_VAL As Double
Dim STDEV_VAL As Double

Dim DW_VAL As Double
Dim TEMP_VAL As Double

Dim TEST_STR As String
Dim LABEL_STR As String

Dim NAME1_STR As String
Dim NAME2_STR As String
Dim NAME3_STR As String
Dim NAME4_STR As String

Dim DATA_ARR() As Double
Dim LAG_ARR() As Double
Dim DIFF_ARR() As Double
Dim LLIKE_VAL As Double
Dim AIC_ARR() As Double

Dim TEMP_GROUP As Variant
Dim TEMP_MATRIX As Variant

Dim OLS_MATRIX As Variant
Dim DATA_VECTOR As Variant
Dim RESIDUALS_VECTOR As Variant

Dim YDATA_VECTOR As Variant
Dim XDATA_MATRIX As Variant

Dim INTERCEPT_FLAG As Boolean
Const PI_VAL As Double = 3.14159265358979

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then
    DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
End If
NOBS = UBound(DATA_VECTOR, 1)

'-------------------------------------------------------------------------------------
'Test for unit root in
Select Case ROOT_MODE
'-------------------------------------------------------------------------------------
Case 0
    ROOT_INT = 0 'level
Case 1
    ROOT_INT = 1 '1st Difference
Case Else
    ROOT_INT = 2 '2nd Difference
'-------------------------------------------------------------------------------------
End Select
'-------------------------------------------------------------------------------------

'-------------------------------------------------------------------------------------
'Include in test equation
Select Case TEST_MODE
'-------------------------------------------------------------------------------------
Case 0 'INTERCEPT_FLAG
    INTERCEPT_FLAG = True
    TREND_VAL = 0
    TEST_STR = "Constant"
    MODEL_INT = 2
Case 1 'TREND_VAL and INTERCEPT_FLAG
    INTERCEPT_FLAG = True
    TREND_VAL = 1
    TEST_STR = "Constand and linear Trend"
    MODEL_INT = 3
Case Else 'None
    INTERCEPT_FLAG = False
    TREND_VAL = 0
    TEST_STR = "None"
    MODEL_INT = 1
'-------------------------------------------------------------------------------------
End Select
'-------------------------------------------------------------------------------------

If NOBS - ROOT_INT < MAX_LAGS Then: GoTo ERROR_LABEL
'Observations must be greater than lags

NOBS = NOBS - ROOT_INT
NAME1_STR = NAME0_STR

'-------------------------------------------------------------------------------------------------
ReDim DATA_ARR(1 To NOBS)
'-------------------------------------------------------------------------------------------------
Select Case ROOT_INT
'If ROOT_INT = 0 Then
Case 0
'-------------------------------------------------------------------------------------------------
    NAME2_STR = NAME0_STR & "(-1)"
    NAME3_STR = ")"
    NAME4_STR = "D(" & NAME1_STR & ")"
    For i = 1 To NOBS: DATA_ARR(i) = DATA_VECTOR(i, 1): Next i
'-------------------------------------------------------------------------------------------------
'ElseIf ROOT_INT = 1 Then
Case 1
'-------------------------------------------------------------------------------------------------
    NAME0_STR = "D(" & NAME1_STR & ")"
    NAME2_STR = "D(" & NAME1_STR & "(-1))"
    NAME3_STR = ",2)"
    NAME4_STR = "D(" & NAME1_STR & ",2)"
    For i = 1 To NOBS: DATA_ARR(i) = DATA_VECTOR(i + 1, 1) - DATA_VECTOR(i, 1): Next i
'-------------------------------------------------------------------------------------------------
'ElseIf ROOT_INT = 2 Then
Case 2
'-------------------------------------------------------------------------------------------------
    NAME0_STR = "D(" & NAME1_STR & ",2)"
    NAME2_STR = "D(" & NAME1_STR & "(-1),2)"
    NAME3_STR = ",3)"
    NAME4_STR = "D(" & NAME1_STR & ",3)"
    For i = 1 To NOBS: DATA_ARR(i) = DATA_VECTOR(i + 2, 1) - 2 * DATA_VECTOR(i + 1, 1) + DATA_VECTOR(i, 1): Next i
'-------------------------------------------------------------------------------------------------
'End If
End Select
'-------------------------------------------------------------------------------------------------
ReDim LAG_ARR(1 To NOBS - 1)
ReDim DIFF_ARR(1 To NOBS - 1)
For i = 1 To NOBS - 1
   LAG_ARR(i) = DATA_ARR(i)
   DIFF_ARR(i) = DATA_ARR(i + 1) - DATA_ARR(i)
Next i
LABEL_STR = "(Fixed)"
'-------------------------------------------------------------------------------------------------
If LAG_AUTOMATIC_FLAG = True Then 'Automatic Selection
'-------------------------------------------------------------------------------------------------
    ReDim AIC_ARR(1 To MAX_LAGS + 1)
    For k = 0 To MAX_LAGS
        GoSub REDIM_LINE
        GoSub EXOGEN_LINE
        OLS_MATRIX = REGRESSION_LS1_FUNC(XDATA_MATRIX, YDATA_VECTOR, INTERCEPT_FLAG, 0, 0)
        LLIKE_VAL = -(NROWS / 2) * Log(2 * PI_VAL) - (NROWS / 2) * Log(OLS_MATRIX(1, 4) / NROWS) - (NROWS / 2)
        '----------------------------------------------------------------------------------------------
        Select Case LAG_LENGTH_MODE
        '----------------------------------------------------------------------------------------------
        Case 0 'Akaike Info Criterion
        '----------------------------------------------------------------------------------------------
            AIC_ARR(k + 1) = -2 * (LLIKE_VAL / NROWS) + ((2 * EXOGEN1_VAL) / NROWS)
            LABEL_STR = "(Automatic Based on AIC, MAXLAG=" & MAX_LAGS & ")"
        '----------------------------------------------------------------------------------------------
        Case 1 'Schwartz Info Criterion
        '----------------------------------------------------------------------------------------------
            AIC_ARR(k + 1) = -2 * (LLIKE_VAL / NROWS) + ((EXOGEN1_VAL * Log(NROWS)) / NROWS)
            LABEL_STR = "(Automatic Based on SIC, MAXLAG=" & MAX_LAGS & ")"
        '----------------------------------------------------------------------------------------------
        Case Else 'Hannan-Quinn Criterion
        '----------------------------------------------------------------------------------------------
            AIC_ARR(k + 1) = -2 * (LLIKE_VAL / NROWS) + ((2 * EXOGEN1_VAL * Log(Log(NROWS))) / NROWS)
            LABEL_STR = "(Automatic Based on HQ, MAXLAG=" & MAX_LAGS & ")"
        '----------------------------------------------------------------------------------------------
        End Select
        '----------------------------------------------------------------------------------------------
    Next k
    TEMP_VAL = AIC_ARR(1)
    For i = 1 To k - 1
        If TEMP_VAL > AIC_ARR(i + 1) Then
            TEMP_VAL = AIC_ARR(i + 1)
            h = i
        End If
    Next i
    k = h
    EXOGEN2_VAL = EXOGEN2_VAL + k
'-------------------------------------------------------------------------------------------------
Else 'User specified
'-------------------------------------------------------------------------------------------------
    k = MAX_LAGS
    GoSub EXOGEN_LINE
'-------------------------------------------------------------------------------------------------
End If
'-------------------------------------------------------------------------------------------------

GoSub RESIDUALS_LINE
GoSub OUTPUT_LINE
ADF_UNIT_ROOT_TEST_FUNC = TEMP_MATRIX

'-------------------------------------------------------------------------------------------------
Exit Function
'-------------------------------------------------------------------------------------------------
REDIM_LINE:
'-------------------------------------------------------------------------------------------------
    NROWS = NOBS - k - 1
    NCOLUMNS = k + 1 + TREND_VAL
    ReDim YDATA_VECTOR(1 To NROWS, 1 To 1)
    ReDim XDATA_MATRIX(1 To NROWS, 1 To NCOLUMNS)
    MEAN_VAL = 0
    For i = 1 To NROWS
       YDATA_VECTOR(i, 1) = DIFF_ARR(i + k)
       MEAN_VAL = MEAN_VAL + YDATA_VECTOR(i, 1)
       XDATA_MATRIX(i, 1) = LAG_ARR(i + k)
       For j = 2 To k + 1: XDATA_MATRIX(i, j) = DIFF_ARR(i + k - j + 1): Next j
       If TREND_VAL = 1 Then: XDATA_MATRIX(i, NCOLUMNS) = i
    Next i
    MEAN_VAL = MEAN_VAL / NROWS
    STDEV_VAL = 0: For i = 1 To NROWS: STDEV_VAL = STDEV_VAL + (YDATA_VECTOR(i, 1) - MEAN_VAL) ^ 2: Next i
    STDEV_VAL = (STDEV_VAL / (NROWS - 1)) ^ 0.5
'-------------------------------------------------------------------------------------------------
Return
'-------------------------------------------------------------------------------------------------
EXOGEN_LINE:
'-------------------------------------------------------------------------------------------------
    If INTERCEPT_FLAG = True Then
        EXOGEN1_VAL = k + 2 + TREND_VAL
        EXOGEN2_VAL = TREND_VAL + 2
    Else
        EXOGEN1_VAL = k + 1 + TREND_VAL
        EXOGEN2_VAL = TREND_VAL + 1
    End If
'-------------------------------------------------------------------------------------------------
Return
'-------------------------------------------------------------------------------------------------
RESIDUALS_LINE:
'-------------------------------------------------------------------------------------------------
    l = 15
    GoSub REDIM_LINE
    If INTERCEPT_FLAG = True Then
        DF_VAL = NROWS - NCOLUMNS - 1 'Residual DF
        If TREND_VAL = 1 Then
            ReDim TEMP_MATRIX(1 To l + k + 9, 1 To 5)
        Else
            ReDim TEMP_MATRIX(1 To l + k + 8, 1 To 5)
        End If
    Else
        DF_VAL = NROWS - NCOLUMNS 'Residual DF
        ReDim TEMP_MATRIX(1 To l + k + 7, 1 To 5)
    End If
    '-------------------------------------------------------------------------------------------------
    TEMP_GROUP = REGRESSION_LS1_FUNC(XDATA_MATRIX, YDATA_VECTOR, INTERCEPT_FLAG, 0, 6)
    OLS_MATRIX = TEMP_GROUP(LBound(TEMP_GROUP) + 0)
    RESIDUALS_VECTOR = TEMP_GROUP(LBound(TEMP_GROUP) + 2)
    Erase TEMP_GROUP
    '-------------------------------------------------------------------------------------------------
    DW_VAL = 0: For i = 2 To NROWS: DW_VAL = DW_VAL + (RESIDUALS_VECTOR(i, 1) - RESIDUALS_VECTOR(i - 1, 1)) ^ 2: Next i
    Erase RESIDUALS_VECTOR
    DW_VAL = DW_VAL / OLS_MATRIX(1, 4)
    LLIKE_VAL = -(NROWS / 2) * Log(2 * PI_VAL) - (NROWS / 2) * Log(OLS_MATRIX(1, 4) / NROWS) - (NROWS / 2)
'-------------------------------------------------------------------------------------------------
Return
'-------------------------------------------------------------------------------------------------
OUTPUT_LINE:
'-------------------------------------------------------------------------------------------------
    For i = 1 To UBound(TEMP_MATRIX, 1): For j = 1 To UBound(TEMP_MATRIX, 2): TEMP_MATRIX(i, j) = "": Next j: Next i
    '-------------------------------------------------------------------------------------------------
    TEMP_MATRIX(1, 1) = "Null Hypothesis: " & NAME0_STR & " has a unit root"
    TEMP_MATRIX(2, 1) = "Exogenous: " & TEST_STR
    TEMP_MATRIX(3, 1) = "Lag Length: " & k & " " & LABEL_STR
    TEMP_MATRIX(4, 4) = "t-Statistic"
    TEMP_MATRIX(4, 5) = "Prob.*"
    
    TEMP_MATRIX(5, 1) = "Augmented Dickey-Fuller test statistic"
    TEMP_MATRIX(6, 1) = "Test critical values: "
    
    TEMP_MATRIX(6, 2) = "1% level"
    TEMP_MATRIX(7, 2) = "5% level"
    TEMP_MATRIX(8, 2) = "10% level"
    
    TEMP_MATRIX(6, 4) = CVADF_FUNC(0.01, NROWS, MODEL_INT)
    TEMP_MATRIX(7, 4) = CVADF_FUNC(0.05, NROWS, MODEL_INT)
    TEMP_MATRIX(8, 4) = CVADF_FUNC(0.1, NROWS, MODEL_INT)
    
    TEMP_MATRIX(9, 1) = "MacKinnon (1996) one-sided p-values."
    TEMP_MATRIX(10, 1) = "Augmented Dickey-Fuller Test Equation"
    TEMP_MATRIX(11, 1) = "Dependent Variable: " & NAME4_STR
    TEMP_MATRIX(12, 1) = "Included observations: " & NROWS & " after adjusting endpoints"
    
    TEMP_MATRIX(14, 1) = "Variable"
    TEMP_MATRIX(14, 2) = "Coefficient"
    TEMP_MATRIX(14, 3) = "Std. Error"
    TEMP_MATRIX(14, 4) = "t-Statistic"
    TEMP_MATRIX(14, 5) = "Prob"
    
    TEMP_MATRIX(15, 1) = NAME2_STR
    
    '-------------------------------------------------------------------------------------------------------
    If INTERCEPT_FLAG = True And TREND_VAL = 0 Then
    '-------------------------------------------------------------------------------------------------------
       j = 7
       For i = k + 1 To 1 Step -1
          h = l + 1 + k - i
          TEMP_MATRIX(h, 2) = OLS_MATRIX(j, 2)
          TEMP_MATRIX(h, 3) = OLS_MATRIX(j, 3)
          j = j + 1
          TEMP_VAL = TEMP_MATRIX(h, 2) / TEMP_MATRIX(h, 3)
          TEMP_MATRIX(h, 4) = TEMP_VAL
          TEMP_MATRIX(h, 5) = 2 * (1 - TDIST_FUNC(Abs(TEMP_VAL), NROWS - (k + 2), True))
       Next i
       j = 6
       h = l + k + 1
       TEMP_MATRIX(h, 2) = OLS_MATRIX(j, 2)
       TEMP_MATRIX(h, 3) = OLS_MATRIX(j, 3)
       TEMP_VAL = TEMP_MATRIX(h, 2) / TEMP_MATRIX(h, 3)
       TEMP_MATRIX(h, 4) = TEMP_VAL
       TEMP_MATRIX(h, 5) = 2 * (1 - TDIST_FUNC(Abs(TEMP_VAL), NROWS - (k + 2), True))
    '-------------------------------------------------------------------------------------------------------
    End If
    '-------------------------------------------------------------------------------------------------------
    
    '-------------------------------------------------------------------------------------------------------
    If TREND_VAL = 1 Then
    '-------------------------------------------------------------------------------------------------------
        j = 7
        For i = k + 2 To 2 Step -1
          h = l + 2 + k - i
          TEMP_MATRIX(h, 2) = OLS_MATRIX(j, 2)
          TEMP_MATRIX(h, 3) = OLS_MATRIX(j, 3)
          j = j + 1
          TEMP_VAL = TEMP_MATRIX(h, 2) / TEMP_MATRIX(h, 3)
          TEMP_MATRIX(h, 4) = TEMP_VAL
          TEMP_MATRIX(h, 5) = 2 * (1 - TDIST_FUNC(Abs(TEMP_VAL), NROWS - (k + 3), True))
        Next i
        j = 6
        h = l + k + 1
        TEMP_MATRIX(h, 2) = OLS_MATRIX(j, 2)
        TEMP_MATRIX(h, 3) = OLS_MATRIX(j, 3)
        TEMP_VAL = TEMP_MATRIX(h, 2) / TEMP_MATRIX(h, 3)
        TEMP_MATRIX(h, 4) = TEMP_VAL
        TEMP_MATRIX(h, 5) = 2 * (1 - TDIST_FUNC(Abs(TEMP_VAL), NROWS - (k + 3), True))
        
        j = 5 + EXOGEN2_VAL
        h = h + 1
        TEMP_MATRIX(h, 2) = OLS_MATRIX(j, 2)
        TEMP_MATRIX(h, 3) = OLS_MATRIX(j, 3)
        TEMP_VAL = TEMP_MATRIX(h, 2) / TEMP_MATRIX(h, 3)
        TEMP_MATRIX(h, 4) = TEMP_VAL
        TEMP_MATRIX(h, 5) = 2 * (1 - TDIST_FUNC(Abs(TEMP_VAL), NROWS - (k + 3), True))
    '-------------------------------------------------------------------------------------------------------
    End If
    '-------------------------------------------------------------------------------------------------------
    
    '-------------------------------------------------------------------------------------------------------
    If INTERCEPT_FLAG = False Then
    '-------------------------------------------------------------------------------------------------------
        j = 6
        For i = k + 1 To 1 Step -1
          h = l + 1 + k - i
          TEMP_MATRIX(h, 2) = OLS_MATRIX(j, 2)
          TEMP_MATRIX(h, 3) = OLS_MATRIX(j, 3)
          j = j + 1
          TEMP_VAL = TEMP_MATRIX(h, 2) / TEMP_MATRIX(h, 3)
          TEMP_MATRIX(h, 4) = TEMP_VAL
          TEMP_MATRIX(h, 5) = 2 * (1 - TDIST_FUNC(Abs(TEMP_VAL), NROWS - (k + 1), True))
        Next i
    '-------------------------------------------------------------------------------------------------------
    End If
    '-------------------------------------------------------------------------------------------------------
    For i = 1 To k: TEMP_MATRIX(l + i, 1) = "D(" & NAME1_STR & "(-" & i & ")" & NAME3_STR: Next i
    '-------------------------------------------------------------------------------------------------------
    TEMP_MATRIX(5, 4) = TEMP_MATRIX(l, 4)
    TEMP_MATRIX(5, 5) = PVADF_FUNC(TEMP_MATRIX(5, 4), NROWS, MODEL_INT)
    
    i = l + k
    If INTERCEPT_FLAG = True Then
        i = i + 1
        TEMP_MATRIX(i, 1) = "C"
    End If
    If TREND_VAL = 1 Then
        i = i + 1
        TEMP_MATRIX(i, 1) = "@Trend"
    End If
    '------------------------------------------------------------------------------------------------------
    
    TEMP_MATRIX(i + 2, 1) = "R-squared"
    TEMP_MATRIX(i + 2, 2) = OLS_MATRIX(4, 4)
    
    TEMP_MATRIX(i + 2, 4) = "Mean dependent var"
    TEMP_MATRIX(i + 2, 5) = MEAN_VAL
    
    'http://en.wikipedia.org/wiki/Coefficient_of_determination
    TEMP_MATRIX(i + 3, 1) = "Adjusted R-squared" 'Perfect
    TEMP_MATRIX(i + 3, 2) = TEMP_MATRIX(i + 2, 2) - (1 - TEMP_MATRIX(i + 2, 2)) * ((NROWS - DF_VAL - 1) / (NROWS - DF_VAL))
    
    TEMP_MATRIX(i + 3, 4) = "S.D. dependent var"
    TEMP_MATRIX(i + 3, 5) = STDEV_VAL
    
    TEMP_MATRIX(i + 4, 1) = "S.E. of regression"
    TEMP_MATRIX(i + 4, 2) = OLS_MATRIX(3, 2) 'same
    
    TEMP_MATRIX(i + 4, 4) = "Akaike info criterion"
    'http://en.wikipedia.org/wiki/Akaike_information_criterion
    TEMP_MATRIX(i + 4, 5) = -2 * (LLIKE_VAL / NROWS) + ((2 * EXOGEN2_VAL) / NROWS)
    
    TEMP_MATRIX(i + 5, 1) = "Sum squared Residuals"
    TEMP_MATRIX(i + 5, 2) = OLS_MATRIX(1, 4) '(5, 2)
    
    TEMP_MATRIX(i + 5, 4) = "Schwarz criterion"
    'http://en.wikipedia.org/wiki/Bayesian_information_criterion
    TEMP_MATRIX(i + 5, 5) = -2 * (LLIKE_VAL / NROWS) + ((EXOGEN2_VAL * Log(NROWS)) / NROWS)
    
    TEMP_MATRIX(i + 6, 1) = "Log likelihood"
    'http://en.wikipedia.org/wiki/Likelihood-ratio_test
    TEMP_MATRIX(i + 6, 2) = LLIKE_VAL
    
    TEMP_MATRIX(i + 6, 4) = "F-statistic"
    TEMP_MATRIX(i + 6, 5) = OLS_MATRIX(4, 2)
    
    TEMP_MATRIX(i + 7, 1) = "Durbin-Watson stat"
    'http://en.wikipedia.org/wiki/Durbin%E2%80%93Watson_statistic
    TEMP_MATRIX(i + 7, 2) = DW_VAL
    
    TEMP_MATRIX(i + 7, 4) = "Prob(F-statistic)" 'Perfect
    TEMP_MATRIX(i + 7, 5) = FDIST_FUNC(TEMP_MATRIX(i + 6, 5), (EXOGEN2_VAL - 1), DF_VAL, True, False)
    
    'The value of df is calculated as follows, when no X columns are removed from the model due
    'to collinearity: if there are k columns of known_x’s and const = TRUE or is omitted,
    'df = n – k – 1. If const = FALSE, df = n - k. In both cases, each X column that was removed
    'due to collinearity increases the value of df by 1.
'------------------------------------------------------------------------------------------------------------------
Return
'------------------------------------------------------------------------------------------------------------------
ERROR_LABEL:
ADF_UNIT_ROOT_TEST_FUNC = Err.number
End Function


Function CVADF_FUNC(ByVal SIZE_VAL As Double, _
ByVal NROWS As Long, _
Optional ByVal MODEL_INT As Long = 1)

'MODEL_INT=1; no constant, no trend
'MODEL_INT=2; constant, no trend
'MODEL_INT=3; constant, trend

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long
Dim m As Long
Dim n As Long
Dim o As Long

Dim ii As Long
Dim jj As Long
Dim kk As Long

Dim MIN_VAL As Long
Dim NPH_VAL As Long

Dim NPTOP_VAL As Long

Dim TOP_VAL As Double
Dim BOT_VAL As Double

Dim QV_VAL As Double
Dim DIFFM_VAL As Double
Dim DIFF_VAL As Double

Dim XOMX_ARR() As Variant
Dim XOMY_ARR() As Variant
Dim BETA_ARR() As Double

Dim NC_MATRIX() As Double
Dim PROB_MATRIX() As Double

Dim CVAL_ARR(1 To 221) As Double
Dim YDATA_VECTOR() As Double
Dim XDATA_MATRIX() As Double

Dim OMEGA_MATRIX() As Variant
Dim INVERSE_MATRIX As Variant

On Error GoTo ERROR_LABEL

If SIZE_VAL < 0 Or SIZE_VAL > 1 Then: GoTo ERROR_LABEL

Select Case MODEL_INT
Case 1 'nc
    NC_MATRIX = READ_NC_DAT_FUNC()
Case 2 'cnt
    NC_MATRIX = READ_CNT_DAT_FUNC()
Case Else '3 'ct
    NC_MATRIX = READ_CT_DAT_FUNC()
End Select
kk = UBound(NC_MATRIX, 2)
PROB_MATRIX = READ_PROB_DAT_FUNC()

DIFFM_VAL = 1000 'calcimin
MIN_VAL = 0
For i = 1 To 221
   DIFF_VAL = Abs(SIZE_VAL - PROB_MATRIX(i, 1))
   If DIFF_VAL < DIFFM_VAL Then
      DIFFM_VAL = DIFF_VAL
      MIN_VAL = i
   End If
Next i

n = 9
NPH_VAL = n / 2
NPTOP_VAL = 221 - NPH_VAL
GoSub EVAL_LINE

'---------------------------------------------------------------------------------------
If MIN_VAL > NPH_VAL And MIN_VAL < NPTOP_VAL Then
'---------------------------------------------------------------------------------------
   m = 4
   ReDim XOMX_ARR(1 To m, 1 To m)
   ReDim XOMY_ARR(1 To m)
   ReDim BETA_ARR(1 To m)
   ReDim OMEGA_MATRIX(1 To n, 1 To n)
   ReDim YDATA_VECTOR(1 To n)
   ReDim XDATA_MATRIX(1 To n, 1 To 4)
   For i = 1 To n
      ii = MIN_VAL - NPH_VAL - 1 + i
      YDATA_VECTOR(i) = CVAL_ARR(ii)
      XDATA_MATRIX(i, 1) = 1
      XDATA_MATRIX(i, 2) = PROB_MATRIX(ii, 2)
      XDATA_MATRIX(i, 3) = PROB_MATRIX(ii, 2) * PROB_MATRIX(ii, 2)
      XDATA_MATRIX(i, 4) = PROB_MATRIX(ii, 2) * PROB_MATRIX(ii, 2) * PROB_MATRIX(ii, 2)
   Next i
   For i = 1 To n
      For j = 1 To n
         ii = MIN_VAL - NPH_VAL - 1 + i
         jj = MIN_VAL - NPH_VAL - 1 + j
         TOP_VAL = PROB_MATRIX(ii, 1) * (1 - PROB_MATRIX(jj, 1))
         BOT_VAL = PROB_MATRIX(jj, 1) * (1 - PROB_MATRIX(ii, 1))
         OMEGA_MATRIX(i, j) = NC_MATRIX(ii, kk) * NC_MATRIX(jj, kk) * Sqr(TOP_VAL / BOT_VAL)
      Next j
   Next i
   For i = 1 To n
      For j = i To n
         OMEGA_MATRIX(j, i) = OMEGA_MATRIX(i, j)
      Next j
   Next i
   GoSub GLS_LINE
   QV_VAL = 0
   For i = 1 To m
      QV_VAL = QV_VAL + BETA_ARR(i) * NORMSINV_FUNC(SIZE_VAL, 0, 1, 0) ^ (i - 1)
   Next i
   CVADF_FUNC = QV_VAL
   
'---------------------------------------------------------------------------------------
ElseIf MIN_VAL < NPH_VAL + 1 Then
'---------------------------------------------------------------------------------------
   m = 4
   o = h + NPH_VAL
   If o < 5 Then o = 5
   n = o
   ReDim XOMX_ARR(1 To m, 1 To m)
   ReDim XOMY_ARR(1 To m)
   ReDim BETA_ARR(1 To m)
   ReDim OMEGA_MATRIX(1 To n, 1 To n)
   ReDim YDATA_VECTOR(1 To n)
   ReDim XDATA_MATRIX(1 To n, 1 To 4)
   For i = 1 To o
      YDATA_VECTOR(i) = CVAL_ARR(i)
      XDATA_MATRIX(i, 1) = 1
      XDATA_MATRIX(i, 2) = PROB_MATRIX(i, 2)
      XDATA_MATRIX(i, 3) = PROB_MATRIX(i, 2) * PROB_MATRIX(i, 2)
      XDATA_MATRIX(i, 4) = PROB_MATRIX(i, 2) * PROB_MATRIX(i, 2) * PROB_MATRIX(i, 2)
   Next i
   For i = 1 To o
      For j = i To o
         TOP_VAL = PROB_MATRIX(i, 1) * (1 - PROB_MATRIX(j, 1))
         BOT_VAL = PROB_MATRIX(j, 1) * (1 - PROB_MATRIX(i, 1))
         OMEGA_MATRIX(i, j) = NC_MATRIX(i, kk) * NC_MATRIX(j, kk) * Sqr(TOP_VAL / BOT_VAL)
      Next j
   Next i
   For i = 1 To o
      For j = i To o
         OMEGA_MATRIX(j, i) = OMEGA_MATRIX(i, j)
      Next j
   Next i
   GoSub GLS_LINE
   QV_VAL = 0
   For i = 1 To m
      QV_VAL = QV_VAL + BETA_ARR(i) * NORMSINV_FUNC(SIZE_VAL, 0, 1, 0) ^ (i - 1)
   Next i
   CVADF_FUNC = QV_VAL

'---------------------------------------------------------------------------------------
Else
'---------------------------------------------------------------------------------------
   o = 221 - MIN_VAL + NPH_VAL
   m = 4
   If o < 5 Then o = 5
   n = o
   ReDim XOMX_ARR(1 To m, 1 To m)
   ReDim XOMY_ARR(1 To m)
   ReDim BETA_ARR(1 To m)
   ReDim OMEGA_MATRIX(1 To n, 1 To n)
   ReDim YDATA_VECTOR(1 To n)
   ReDim XDATA_MATRIX(1 To n, 1 To 4)
   For i = 1 To o
      ii = 222 - i
      YDATA_VECTOR(i) = CVAL_ARR(ii)
      XDATA_MATRIX(i, 1) = 1
      XDATA_MATRIX(i, 2) = PROB_MATRIX(ii, 2)
      XDATA_MATRIX(i, 3) = PROB_MATRIX(ii, 2) * PROB_MATRIX(ii, 2)
      XDATA_MATRIX(i, 4) = PROB_MATRIX(ii, 2) * PROB_MATRIX(ii, 2) * PROB_MATRIX(ii, 2)
   Next i
   For i = 1 To o
      For j = i To o
         OMEGA_MATRIX(i, j) = 0
         If i = j Then OMEGA_MATRIX(i, j) = 1
      Next j
   Next i
   For i = 1 To o
      For j = i To o
         OMEGA_MATRIX(j, i) = OMEGA_MATRIX(i, j)
      Next j
   Next i
   GoSub GLS_LINE
   QV_VAL = 0
   For i = 1 To m
      QV_VAL = QV_VAL + BETA_ARR(i) * NORMSINV_FUNC(SIZE_VAL, 0, 1, 0) ^ (i - 1)
   Next i
   CVADF_FUNC = QV_VAL

'---------------------------------------------------------------------------------------
End If
'---------------------------------------------------------------------------------------

Exit Function
'---------------------------------------------------------------------------------------
EVAL_LINE:
'---------------------------------------------------------------------------------------
   If NROWS = 0 Then
      For i = 1 To 221
         CVAL_ARR(i) = NC_MATRIX(i, 1)
      Next i
   Else
      If MODEL_INT = 1 Or MODEL_INT = 2 Then
         For i = 1 To 221
            CVAL_ARR(i) = NC_MATRIX(i, 1) + NC_MATRIX(i, 2) / NROWS + NC_MATRIX(i, 3) / (NROWS * NROWS)
         Next i
      Else
         For i = 1 To 221
            CVAL_ARR(i) = NC_MATRIX(i, 1) + NC_MATRIX(i, 2) / NROWS + NC_MATRIX(i, 3) / (NROWS * NROWS) + NC_MATRIX(i, 4) / (NROWS * NROWS * NROWS)
         Next i
      End If
   End If
'---------------------------------------------------------------------------------------
Return
'---------------------------------------------------------------------------------------
GLS_LINE:
'---------------------------------------------------------------------------------------
   INVERSE_MATRIX = MATRIX_INVERSE_FUNC(OMEGA_MATRIX, 0)
   For j = 1 To m
      XOMY_ARR(j) = 0
      For l = j To m
         XOMX_ARR(j, l) = 0
      Next l
   Next j
   
   For i = 1 To n
      For k = 1 To n
         For j = 1 To m
            XOMY_ARR(j) = XOMY_ARR(j) + XDATA_MATRIX(i, j) * INVERSE_MATRIX(k, i) * YDATA_VECTOR(k)
            For l = j To m
               XOMX_ARR(j, l) = XOMX_ARR(j, l) + XDATA_MATRIX(i, j) * INVERSE_MATRIX(k, i) * XDATA_MATRIX(k, l)
            Next l
         Next j
      Next k
   Next i
   For j = 1 To m
      For l = j To m
         XOMX_ARR(l, j) = XOMX_ARR(j, l)
      Next l
   Next j
   INVERSE_MATRIX = MATRIX_INVERSE_FUNC(XOMX_ARR, 0)
   For i = 1 To m
      BETA_ARR(i) = 0
      For j = 1 To m
         BETA_ARR(i) = BETA_ARR(i) + INVERSE_MATRIX(i, j) * XOMY_ARR(j)
      Next j
   Next i
'---------------------------------------------------------------------------------------
Return
'---------------------------------------------------------------------------------------
ERROR_LABEL:
CVADF_FUNC = Err.number
End Function

Function PVADF_FUNC(ByVal TSTAT_VAL As Double, _
ByVal NROWS As Long, _
Optional ByVal MODEL_INT As Long = 1)

'MODEL_INT=1; no constant, no trend
'MODEL_INT=2; constant, no trend
'MODEL_INT=3; constant, trend

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long
Dim m As Long
Dim n As Long
Dim o As Long

Dim ii As Long
Dim jj As Long
Dim kk As Long

Dim MIN_VAL As Long
Dim NPH_VAL As Long
Dim NPTOP_VAL As Long

Dim TOP_VAL As Double
Dim BOT_VAL As Double
Dim QV_VAL As Double

Dim DIFFM_VAL As Double
Dim DIFF_VAL As Double

Dim XOMX_ARR() As Variant
Dim XOMY_ARR() As Variant
Dim BETA_ARR() As Double

Dim NC_MATRIX() As Double
Dim PROB_MATRIX() As Double

Dim CVAL_ARR(1 To 221) As Double
Dim YDATA_VECTOR() As Double
Dim XDATA_MATRIX() As Double
Dim OMEGA_MATRIX() As Variant
Dim INVERSE_MATRIX As Variant

On Error GoTo ERROR_LABEL

Select Case MODEL_INT
Case 1 'nc
    NC_MATRIX = READ_NC_DAT_FUNC()
Case 2 'cnt
    NC_MATRIX = READ_CNT_DAT_FUNC()
Case Else '3 'ct
    NC_MATRIX = READ_CT_DAT_FUNC()
End Select
kk = UBound(NC_MATRIX, 2)
PROB_MATRIX = READ_PROB_DAT_FUNC()
GoSub EVAL_LINE

'calciminpv
DIFFM_VAL = 1000
MIN_VAL = 0
For i = 1 To 221
   DIFF_VAL = Abs(TSTAT_VAL - CVAL_ARR(i))
   If DIFF_VAL < DIFFM_VAL Then
      DIFFM_VAL = DIFF_VAL
      h = i
   End If
Next i
'-------------------------------------------------------------------------

n = 9
NPH_VAL = n / 2
NPTOP_VAL = 221 - NPH_VAL
'---------------------------------------------------------------------------------------
If h > NPH_VAL And h < NPTOP_VAL Then
'---------------------------------------------------------------------------------------
   m = 4
   n = 9
   ReDim XOMX_ARR(1 To m, 1 To m)
   ReDim XOMY_ARR(1 To m)
   ReDim BETA_ARR(1 To m)
   ReDim OMEGA_MATRIX(1 To n, 1 To n)
   ReDim YDATA_VECTOR(1 To n)
   ReDim XDATA_MATRIX(1 To n, 1 To 4)
   For i = 1 To n
      ii = h - NPH_VAL - 1 + i
      YDATA_VECTOR(i) = PROB_MATRIX(ii, 2)
      XDATA_MATRIX(i, 1) = 1
      XDATA_MATRIX(i, 2) = CVAL_ARR(ii)
      XDATA_MATRIX(i, 3) = CVAL_ARR(ii) * CVAL_ARR(ii)
      XDATA_MATRIX(i, 4) = CVAL_ARR(ii) * CVAL_ARR(ii) * CVAL_ARR(ii)
   Next i
   For i = 1 To n
      For j = 1 To n
         ii = h - NPH_VAL - 1 + i
         jj = h - NPH_VAL - 1 + j
         TOP_VAL = PROB_MATRIX(ii, 1) * (1 - PROB_MATRIX(jj, 1))
         BOT_VAL = PROB_MATRIX(jj, 1) * (1 - PROB_MATRIX(ii, 1))
         OMEGA_MATRIX(i, j) = NC_MATRIX(ii, kk) * NC_MATRIX(jj, kk) * Sqr(TOP_VAL / BOT_VAL)
      Next j
   Next i
   For i = 1 To n
      For j = i To n
         OMEGA_MATRIX(j, i) = OMEGA_MATRIX(i, j)
      Next j
   Next i
   GoSub GLS_LINE
   QV_VAL = 0
   For i = 1 To m
      QV_VAL = QV_VAL + BETA_ARR(i) * TSTAT_VAL ^ (i - 1)
   Next i
   PVADF_FUNC = NORMSDIST_FUNC(QV_VAL, 0, 1, 0)
'---------------------------------------------------------------------------------------
ElseIf h < NPH_VAL + 1 Then
'---------------------------------------------------------------------------------------
   m = 4
   o = h + NPH_VAL
   If o < 5 Then o = 5
   n = o
   ReDim XOMX_ARR(1 To m, 1 To m)
   ReDim XOMY_ARR(1 To m)
   ReDim BETA_ARR(1 To m)
   ReDim OMEGA_MATRIX(1 To n, 1 To n)
   ReDim YDATA_VECTOR(1 To n)
   ReDim XDATA_MATRIX(1 To n, 1 To 4)
   For i = 1 To o
      YDATA_VECTOR(i) = PROB_MATRIX(i, 2)
      XDATA_MATRIX(i, 1) = 1
      XDATA_MATRIX(i, 2) = CVAL_ARR(i)
      XDATA_MATRIX(i, 3) = CVAL_ARR(i) * CVAL_ARR(i)
      XDATA_MATRIX(i, 4) = CVAL_ARR(i) * CVAL_ARR(i) * CVAL_ARR(i)
   Next i
   For i = 1 To o
      For j = i To o
         TOP_VAL = PROB_MATRIX(i, 1) * (1 - PROB_MATRIX(j, 1))
         BOT_VAL = PROB_MATRIX(j, 1) * (1 - PROB_MATRIX(i, 1))
         OMEGA_MATRIX(i, j) = NC_MATRIX(i, kk) * NC_MATRIX(j, kk) * Sqr(TOP_VAL / BOT_VAL)
      Next j
   Next i
   For i = 1 To o
      For j = i To o
         OMEGA_MATRIX(j, i) = OMEGA_MATRIX(i, j)
      Next j
   Next i
   GoSub GLS_LINE
   QV_VAL = 0
   For i = 1 To m
      QV_VAL = QV_VAL + BETA_ARR(i) * TSTAT_VAL ^ (i - 1)
   Next i
   PVADF_FUNC = NORMSDIST_FUNC(QV_VAL, 0, 1, 0)
   If h = 1 And PVADF_FUNC > PROB_MATRIX(1, 1) Then PVADF_FUNC = PROB_MATRIX(1, 1)
'---------------------------------------------------------------------------------------
Else
'---------------------------------------------------------------------------------------
   o = 221 - h + NPH_VAL
   m = 4
   If o < 5 Then o = 5
   n = o
   ReDim XOMX_ARR(1 To m, 1 To m)
   ReDim XOMY_ARR(1 To m)
   ReDim BETA_ARR(1 To m)
   ReDim OMEGA_MATRIX(1 To n, 1 To n)
   ReDim YDATA_VECTOR(1 To n)
   ReDim XDATA_MATRIX(1 To n, 1 To 4)
   For i = 1 To o
      ii = 222 - i
      YDATA_VECTOR(i) = PROB_MATRIX(ii, 2)
      XDATA_MATRIX(i, 1) = 1
      XDATA_MATRIX(i, 2) = CVAL_ARR(ii)
      XDATA_MATRIX(i, 3) = CVAL_ARR(ii) * CVAL_ARR(ii)
      XDATA_MATRIX(i, 4) = CVAL_ARR(ii) * CVAL_ARR(ii) * CVAL_ARR(ii)
   Next i
   For i = 1 To o
      For j = i To o
         OMEGA_MATRIX(i, j) = 0
         If i = j Then OMEGA_MATRIX(i, j) = 1
      Next j
   Next i
   For i = 1 To o
      For j = i To o
         OMEGA_MATRIX(j, i) = OMEGA_MATRIX(i, j)
      Next j
   Next i
   GoSub GLS_LINE
   QV_VAL = 0
   For i = 1 To m
      QV_VAL = QV_VAL + BETA_ARR(i) * TSTAT_VAL ^ (i - 1)
   Next i
   PVADF_FUNC = NORMSDIST_FUNC(QV_VAL, 0, 1, 0)
   If h = 221 And PVADF_FUNC < PROB_MATRIX(221, 1) Then PVADF_FUNC = PROB_MATRIX(221, 1)
'---------------------------------------------------------------------------------------
End If
'---------------------------------------------------------------------------------------
   
Exit Function
'---------------------------------------------------------------------------------------
EVAL_LINE:
'---------------------------------------------------------------------------------------
   If NROWS = 0 Then
      For i = 1 To 221
         CVAL_ARR(i) = NC_MATRIX(i, 1)
      Next i
   Else
      If MODEL_INT = 1 Or MODEL_INT = 2 Then
         For i = 1 To 221
            CVAL_ARR(i) = NC_MATRIX(i, 1) + NC_MATRIX(i, 2) / NROWS + NC_MATRIX(i, 3) / (NROWS * NROWS)
         Next i
      Else
         For i = 1 To 221
            CVAL_ARR(i) = NC_MATRIX(i, 1) + NC_MATRIX(i, 2) / NROWS + NC_MATRIX(i, 3) / (NROWS * NROWS) + NC_MATRIX(i, 4) / (NROWS * NROWS * NROWS)
         Next i
      End If
   End If
'---------------------------------------------------------------------------------------
Return
'---------------------------------------------------------------------------------------
GLS_LINE:
'---------------------------------------------------------------------------------------
   INVERSE_MATRIX = MATRIX_INVERSE_FUNC(OMEGA_MATRIX, 0)
   For j = 1 To m
      XOMY_ARR(j) = 0
      For l = j To m
         XOMX_ARR(j, l) = 0
      Next l
   Next j
   For i = 1 To n
      For k = 1 To n
         For j = 1 To m
            XOMY_ARR(j) = XOMY_ARR(j) + XDATA_MATRIX(i, j) * INVERSE_MATRIX(k, i) * YDATA_VECTOR(k)
            For l = j To m
               XOMX_ARR(j, l) = XOMX_ARR(j, l) + XDATA_MATRIX(i, j) * INVERSE_MATRIX(k, i) * XDATA_MATRIX(k, l)
            Next l
         Next j
      Next k
   Next i
   For j = 1 To m
      For l = j To m
         XOMX_ARR(l, j) = XOMX_ARR(j, l)
      Next l
   Next j
   INVERSE_MATRIX = MATRIX_INVERSE_FUNC(XOMX_ARR, 0)
   For i = 1 To m
      BETA_ARR(i) = 0
      For j = 1 To m
         BETA_ARR(i) = BETA_ARR(i) + INVERSE_MATRIX(i, j) * XOMY_ARR(j)
      Next j
   Next i
'---------------------------------------------------------------------------------------
Return
'---------------------------------------------------------------------------------------
ERROR_LABEL:
PVADF_FUNC = Err.number
End Function


Private Function READ_CNT_DAT_FUNC() 'Perfect

Dim i As Long
Dim j As Long

Dim ii As Long
Dim jj As Long

Const DELIM_CHR As String = ","
Const TAB_CHR As String = "|"

Const NROWS As Long = 221
Const NCOLUMNS As Long = 4

Dim TEMP_STR As String
Dim LINE_STR As String

On Error GoTo ERROR_LABEL

ReDim CNT_MATRIX(1 To NROWS, 1 To NCOLUMNS) As Double

LINE_STR = _
"-4.64987370E+00,-1.95851280E+01,-1.34048590E+02,2.89875990E-03|-4.49316480E+00,-1.68079260E+01,-1.22047100E+02,2.20830000E-03|-4.26756480E+00,-1.44258310E+01,-8.76442970E+01,1.41402470E-03|-4.09121760E+00,-1.24413670E+01,-6.84923330E+01,1.08826950E-03|-3.90583470E+00,-1.05131020E+01,-5.19184280E+01,7.79588300E-04|-3.79271540E+00,-9.40902720E+00,-4.42084420E+01,6.63188640E-04|-3.71001050E+00,-8.67809360E+00,-3.78308660E+01,5.94975410E-04|-3.64357500E+00,-8.18220490E+00,-3.23557670E+01,5.50074570E-04|-3.58908100E+00,-7.67292770E+00,-3.04219750E+01,5.09673130E-04|-3.54215560E+00,-7.31195780E+00,-2.73195670E+01,4.84843890E-04|-3.50073510E+00,-6.99994490E+00,-2.50381040E+01,4.52856030E-04|-3.46378470E+00,-6.70242390E+00,-2.38331550E+01,4.32121990E-04|-3.43018550E+00,-6.46018350E+00,-2.21506220E+01,4.06255300E-04|-3.29800440E+00,-5.50417610E+00,-1.70795870E+01,3.53165460E-04|-3.20034750E+00,-4.81849350E+00,-1.45186990E+01,3.11442630E-04|-3.12202150E+00,-4.31564080E+00," & _
"-1.23633870E+01,2.84337420E-04|-3.05604840E+00,-3.93601260E+00,-1.03559880E+01,2.73653170E-04|-2.99890100E+00,-3.60044950E+00,-9.29292320E+00,2.56073190E-04|-2.94847150E+00,-3.31013540E+00,-8.28659120E+00,2.44197270E-04|-2.90288070E+00,-3.06896080E+00,-7.37533010E+00,2.37989360E-04|-2.86137940E+00,-2.85838900E+00,-6.53569370E+00,2.25682900E-04|-2.82317070E+00,-2.67048670E+00,-5.86840310E+00,2.20937510E-04|-2.78768100E+00,-2.50095840E+00,-5.27325270E+00,2.13809440E-04|-2.75458890E+00,-2.34149250E+00,-4.85796830E+00,2.09083480E-04|-2.72349400E+00,-2.19507030E+00,-4.50796210E+00,2.04127340E-04|"

LINE_STR = LINE_STR & _
"-2.69403140E+00,-2.07205780E+00,-4.01703280E+00,1.99627890E-04|-2.66616730E+00,-1.95208490E+00,-3.67143420E+00,1.94010440E-04|-2.63957870E+00,-1.84559890E+00,-3.24005360E+00,1.93347730E-04|-2.61420840E+00,-1.74124630E+00,-2.97563010E+00,1.88540580E-04|-2.58996900E+00,-1.63820110E+00,-2.81890500E+00,1.83426140E-04|-2.56668520E+00,-1.54070580E+00,-2.68247380E+00,1.78677180E-04|-2.54429920E+00,-1.44559090E+00,-2.61412520E+00,1.75296190E-04|-2.52265290E+00,-1.36364720E+00,-2.43311780E+00,1.71985770E-04|-2.50177340E+00,-1.28604560E+00,-2.23765680E+00,1.67101550E-04|-2.48150820E+00,-1.21465600E+00,-2.00286700E+00,1.62090840E-04|-2.46181270E+00,-1.14383670E+00,-1.94105830E+00,1.58183200E-04|-2.44273980E+00,-1.07958150E+00,-1.73798970E+00,1.57070670E-04|-2.42423730E+00,-1.01298710E+00,-1.61572650E+00,1.55400080E-04|-2.40612350E+00,-9.58630040E-01,-1.35841820E+00,1.54969570E-04|-2.38847050E+00,-9.03178550E-01,-1.16452860E+00,1.53807260E-04|-2.37128340E+00,-8.48970900E-01," & _
"-1.03619530E+00,1.51974440E-04|-2.35445880E+00,-7.98714090E-01,-8.91038680E-01,1.51552350E-04|-2.33803930E+00,-7.47088500E-01,-8.21825540E-01,1.50566030E-04|-2.32195150E+00,-7.01559670E-01,-6.37969280E-01,1.49281470E-04|-2.30617610E+00,-6.59422930E-01,-4.33144740E-01,1.48307060E-04|-2.29075400E+00,-6.12859100E-01,-3.60485290E-01,1.47211400E-04|-2.27562970E+00,-5.67799430E-01,-2.91214910E-01,1.47191670E-04|-2.26078800E+00,-5.24092360E-01,-2.56068390E-01,1.45802970E-04|-2.24625750E+00,-4.76097230E-01,-3.09948310E-01,1.45286490E-04|-2.23185660E+00,-4.39292920E-01,-2.24408050E-01,1.45059190E-04|"

LINE_STR = LINE_STR & _
"-2.21777300E+00,-3.97015360E-01,-2.48563180E-01,1.41886390E-04|-2.20383520E+00,-3.64448510E-01,-1.13545540E-01,1.41716400E-04|-2.19019140E+00,-3.25323520E-01,-9.32167840E-02,1.42182420E-04|-2.17673910E+00,-2.87471140E-01,-1.03100170E-01,1.41556110E-04|-2.16348830E+00,-2.50166620E-01,-1.06751310E-01,1.40502260E-04|-2.15038740E+00,-2.16012050E-01,-1.04766510E-01,1.40150330E-04|-2.13744320E+00,-1.85833600E-01,-5.12113170E-02,1.40113180E-04|-2.12464200E+00,-1.56834310E-01,-2.92048740E-02,1.38833780E-04|-2.11200880E+00,-1.32133460E-01,1.28528070E-01,1.38577940E-04|-2.09954740E+00,-9.98158910E-02,9.90583410E-02,1.37427970E-04|-2.08722950E+00,-7.03784620E-02,1.08004150E-01,1.37938840E-04|-2.07506090E+00,-4.20448890E-02,1.41012160E-01,1.36219270E-04|-2.06303520E+00,-1.51416840E-02,1.88653040E-01,1.36393560E-04|-2.05113600E+00,1.59520550E-02,1.00126260E-01,1.34900150E-04|-2.03939180E+00,4.45225290E-02,8.65651460E-02,1.34727760E-04|-2.02770770E+00,6.70592340E-02,1.68190540E-01," & _
"1.34076480E-04|-2.01614900E+00,9.09048060E-02,1.95508290E-01,1.34093390E-04|-2.00469200E+00,1.16653230E-01,1.57044510E-01,1.32916550E-04|-1.99332760E+00,1.39687180E-01,1.72478840E-01,1.31511970E-04|-1.98206190E+00,1.61731570E-01,1.98988190E-01,1.31624020E-04|-1.97088600E+00,1.82374810E-01,2.52360520E-01,1.30619070E-04|-1.95982920E+00,2.03915250E-01,2.87272090E-01,1.30187640E-04|-1.94882230E+00,2.23721340E-01,3.28128560E-01,1.29905730E-04|-1.93792660E+00,2.47748430E-01,2.61253130E-01,1.29620200E-04|-1.92708530E+00,2.67457490E-01,2.69793700E-01,1.27870680E-04|"

LINE_STR = LINE_STR & _
"-1.91631510E+00,2.86731530E-01,2.87675810E-01,1.28196500E-04|-1.90562740E+00,3.04802710E-01,3.14403680E-01,1.27501120E-04|-1.89500920E+00,3.24877400E-01,2.92892800E-01,1.26334210E-04|-1.88443730E+00,3.41833010E-01,3.30949590E-01,1.26171400E-04|-1.87393960E+00,3.60933340E-01,2.94157690E-01,1.25421240E-04|-1.86352430E+00,3.81420670E-01,2.26145520E-01,1.24905660E-04|-1.85313020E+00,3.97161050E-01,2.43030660E-01,1.24870500E-04|-1.84279870E+00,4.09604300E-01,3.61118230E-01,1.24508760E-04|-1.83254900E+00,4.28278580E-01,3.50446680E-01,1.24295660E-04|-1.82231200E+00,4.43125330E-01,3.83730890E-01,1.23392860E-04|-1.81214760E+00,4.58253950E-01,4.17360170E-01,1.23545840E-04|-1.80201890E+00,4.74561760E-01,3.97805680E-01,1.22993780E-04|-1.79194410E+00,4.91566070E-01,3.71603520E-01,1.21153980E-04|-1.78189270E+00,5.06697550E-01,3.59312580E-01,1.20903620E-04|-1.77185470E+00,5.20200780E-01,3.68527900E-01,1.19909790E-04|-1.76187840E+00,5.35521310E-01,3.73999270E-01,1.20960960E-04|" & _
"-1.75190350E+00,5.46697200E-01,4.29956180E-01,1.20960990E-04|-1.74198680E+00,5.61808860E-01,4.09029970E-01,1.21276210E-04|-1.73210820E+00,5.74674690E-01,4.28680530E-01,1.20350100E-04|-1.72223750E+00,5.87088260E-01,4.51750840E-01,1.20436510E-04|-1.71237980E+00,5.99176970E-01,4.88603030E-01,1.20512060E-04|-1.70255090E+00,6.11706710E-01,4.94742750E-01,1.20711790E-04|-1.69274080E+00,6.25060260E-01,4.60001670E-01,1.20505940E-04|-1.68296930E+00,6.36590230E-01,4.95645930E-01,1.21111630E-04|-1.67319390E+00,6.48056540E-01,5.08040980E-01,1.20385830E-04|"

LINE_STR = LINE_STR & _
"-1.66340910E+00,6.56720360E-01,5.56076080E-01,1.21025410E-04|-1.65364970E+00,6.67693150E-01,5.59114980E-01,1.21493890E-04|-1.64390620E+00,6.77427990E-01,5.90762770E-01,1.22617310E-04|-1.63418860E+00,6.89949940E-01,5.68256470E-01,1.22131190E-04|-1.62446870E+00,7.01008340E-01,5.72407050E-01,1.22939420E-04|-1.61473640E+00,7.10445650E-01,6.00846760E-01,1.21590890E-04|-1.60501320E+00,7.20471810E-01,6.14998290E-01,1.21139440E-04|-1.59527550E+00,7.29232890E-01,6.50890300E-01,1.21202660E-04|-1.58554070E+00,7.39468900E-01,6.45162600E-01,1.20821450E-04|-1.57583400E+00,7.49174100E-01,6.42063030E-01,1.20867090E-04|-1.56612260E+00,7.61231290E-01,6.01418330E-01,1.21783660E-04|-1.55637250E+00,7.68837280E-01,6.46112020E-01,1.22075160E-04|-1.54662620E+00,7.77495220E-01,6.71642320E-01,1.23187910E-04|-1.53688870E+00,7.90448020E-01,5.99713090E-01,1.23469500E-04|-1.52711930E+00,8.00445310E-01,6.06133390E-01,1.24518860E-04|-1.51735710E+00,8.09419870E-01,6.17213200E-01,1.23847410E-04|" & _
"-1.50758290E+00,8.20114120E-01,5.93234850E-01,1.23822380E-04|-1.49776720E+00,8.30442740E-01,5.87129980E-01,1.23802600E-04|-1.48791220E+00,8.39049730E-01,5.90876870E-01,1.23079830E-04|-1.47804990E+00,8.47326740E-01,6.16080620E-01,1.22666400E-04|-1.46816190E+00,8.56771500E-01,6.09798420E-01,1.22522230E-04|-1.45824630E+00,8.64358030E-01,6.49883850E-01,1.23309110E-04|-1.44828830E+00,8.74517740E-01,6.15436660E-01,1.23472160E-04|-1.43828130E+00,8.82631130E-01,6.29739920E-01,1.23805570E-04|-1.42825360E+00,8.89563980E-01,6.74681720E-01,1.24178420E-04|"

LINE_STR = LINE_STR & _
"-1.41820080E+00,8.99290490E-01,6.74287930E-01,1.24229410E-04|-1.40808850E+00,9.08299420E-01,6.71598190E-01,1.25063860E-04|-1.39793420E+00,9.17038480E-01,6.68775960E-01,1.26252260E-04|-1.38776940E+00,9.28590170E-01,6.19950670E-01,1.26993550E-04|-1.37752300E+00,9.36042890E-01,6.38466490E-01,1.26660640E-04|-1.36720570E+00,9.44268470E-01,6.25115100E-01,1.27371160E-04|-1.35683390E+00,9.51881270E-01,6.37668790E-01,1.27217560E-04|-1.34640790E+00,9.59590070E-01,6.61272870E-01,1.29444530E-04|-1.33590580E+00,9.63665450E-01,7.45781170E-01,1.30002000E-04|-1.32536850E+00,9.71104470E-01,7.96431930E-01,1.30079920E-04|-1.31476240E+00,9.78554050E-01,8.32130550E-01,1.30425070E-04|-1.30411060E+00,9.88356920E-01,8.19115050E-01,1.31390370E-04|-1.29337830E+00,9.96972870E-01,8.68083120E-01,1.31811010E-04|-1.28254970E+00,1.00473370E+00,8.97575640E-01,1.32048240E-04|-1.27166060E+00,1.01463990E+00,9.08144940E-01,1.32579590E-04|-1.26069750E+00,1.02589370E+00,8.76599850E-01,1.33324930E-04|" & _
"-1.24962500E+00,1.03491090E+00,8.88964560E-01,1.33327580E-04|-1.23847250E+00,1.04383860E+00,9.18009850E-01,1.33987110E-04|-1.22718530E+00,1.05342070E+00,9.05953800E-01,1.34663300E-04|-1.21580710E+00,1.05979950E+00,1.00291030E+00,1.34813240E-04|-1.20433420E+00,1.06940940E+00,1.01052870E+00,1.35852210E-04|-1.19273740E+00,1.08170230E+00,9.65153250E-01,1.36348580E-04|-1.18104710E+00,1.09169380E+00,1.00855060E+00,1.36946560E-04|-1.16922360E+00,1.10262910E+00,1.02685100E+00,1.37888060E-04|-1.15725650E+00,1.11138780E+00,1.06300200E+00,1.38835350E-04|"

LINE_STR = LINE_STR & _
"-1.14518710E+00,1.12727070E+00,9.73421830E-01,1.39318660E-04|-1.13295540E+00,1.13902150E+00,9.75719970E-01,1.40415490E-04|-1.12055680E+00,1.14880080E+00,1.02727980E+00,1.40926500E-04|-1.10802510E+00,1.16215140E+00,9.98114800E-01,1.42205210E-04|-1.09536890E+00,1.17514720E+00,9.99092080E-01,1.42298170E-04|-1.08250270E+00,1.18427600E+00,1.10297910E+00,1.44247660E-04|-1.06944710E+00,1.19554970E+00,1.15152050E+00,1.46354270E-04|-1.05623960E+00,1.20734690E+00,1.19217590E+00,1.47618550E-04|-1.04285620E+00,1.22323900E+00,1.13465130E+00,1.48520010E-04|-1.02926200E+00,1.23480750E+00,1.19450830E+00,1.49028070E-04|-1.01546350E+00,1.24338790E+00,1.32529130E+00,1.49478630E-04|-1.00143750E+00,1.25678310E+00,1.35600220E+00,1.52081540E-04|-9.87223610E-01,1.27294850E+00,1.36236590E+00,1.52963260E-04|-9.72765250E-01,1.28498430E+00,1.44983760E+00,1.53746000E-04|-9.58049680E-01,1.29908450E+00,1.46499260E+00,1.54483520E-04|-9.43077890E-01,1.31167120E+00,1.51992570E+00,1.54024270E-04|" & _
"-9.27881470E-01,1.32712470E+00,1.53352710E+00,1.54434620E-04|-9.12400620E-01,1.34580790E+00,1.46857480E+00,1.56326750E-04|-8.96623630E-01,1.36034650E+00,1.51406670E+00,1.56852300E-04|-8.80514390E-01,1.37072040E+00,1.61707010E+00,1.56741020E-04|-8.64152130E-01,1.39055300E+00,1.52592640E+00,1.57605270E-04|-8.47412710E-01,1.40463440E+00,1.55291400E+00,1.59899490E-04|-8.30346170E-01,1.41738850E+00,1.62092500E+00,1.63509120E-04|-8.12944780E-01,1.43622410E+00,1.55519200E+00,1.65864880E-04|-7.95148640E-01,1.45111810E+00,1.55628240E+00,1.65908670E-04|"

LINE_STR = LINE_STR & _
"-7.76966610E-01,1.46770190E+00,1.52495730E+00,1.66384010E-04|-7.58433350E-01,1.48159320E+00,1.56803240E+00,1.66629670E-04|-7.39489360E-01,1.50277830E+00,1.40167700E+00,1.70627590E-04|-7.20017650E-01,1.51550650E+00,1.42024730E+00,1.74211950E-04|-7.00079600E-01,1.53081870E+00,1.37415960E+00,1.75938260E-04|-6.79656340E-01,1.54329390E+00,1.40056420E+00,1.79418200E-04|-6.58748340E-01,1.56282660E+00,1.27237230E+00,1.82150450E-04|-6.37282880E-01,1.57678630E+00,1.29496840E+00,1.84588660E-04|-6.15200900E-01,1.59437280E+00,1.21096360E+00,1.88364830E-04|-5.92491850E-01,1.61248360E+00,1.10007070E+00,1.90578830E-04|-5.69038370E-01,1.62110590E+00,1.15530000E+00,1.94338410E-04|-5.45036020E-01,1.64008160E+00,1.11980990E+00,1.94891170E-04|-5.20142090E-01,1.64793580E+00,1.29046790E+00,1.97870260E-04|-4.94482610E-01,1.66632490E+00,1.23304000E+00,2.01233360E-04|-4.67827510E-01,1.67269210E+00,1.41638460E+00,2.02005460E-04|-4.40226520E-01,1.69126180E+00,1.29208160E+00,2.06169490E-04|" & _
"-4.11509140E-01,1.70306790E+00,1.39560810E+00,2.07911460E-04|-3.81639600E-01,1.72183230E+00,1.28862760E+00,2.11439490E-04|-3.50523280E-01,1.74493350E+00,1.24448790E+00,2.16080960E-04|-3.17849630E-01,1.75839880E+00,1.33093490E+00,2.21376160E-04|-2.83586690E-01,1.78196710E+00,1.19394490E+00,2.28246100E-04|-2.47529360E-01,1.80434320E+00,1.14925540E+00,2.33442800E-04|-2.09214340E-01,1.81286130E+00,1.46813720E+00,2.39124570E-04|-1.68548430E-01,1.82915240E+00,1.68110820E+00,2.42962010E-04|-1.25067700E-01,1.84842290E+00,1.84481040E+00,2.49561330E-04|"

LINE_STR = LINE_STR & _
"-7.84523120E-02,1.89359480E+00,1.64860300E+00,2.55395640E-04|-2.77989500E-02,1.92679520E+00,1.70412550E+00,2.58295140E-04|2.77063920E-02,1.96193690E+00,1.89497380E+00,2.68973170E-04|8.92866620E-02,1.99572090E+00,2.20391150E+00,2.82669730E-04|1.58561160E-01,2.03916920E+00,2.52970690E+00,3.00694150E-04|2.38244090E-01,2.09093620E+00,2.97927990E+00,3.23332740E-04|3.32707960E-01,2.13301790E+00,4.31197300E+00,3.52438730E-04|4.49585240E-01,2.28005210E+00,4.21149180E+00,3.91832600E-04|6.07142020E-01,2.45783500E+00,5.46525850E+00,4.42404280E-04|6.46625610E-01,2.52449550E+00,5.59035440E+00,4.61507710E-04|6.90399360E-01,2.58339490E+00,5.92619450E+00,4.82376410E-04|7.38979230E-01,2.65411950E+00,6.36089760E+00,5.04149020E-04|7.93956210E-01,2.73381030E+00,7.13523790E+00,5.46706770E-04|8.57903850E-01,2.85128740E+00,8.08399360E+00,5.98704110E-04|9.34044810E-01,3.04963550E+00,8.20898040E+00,6.61945750E-04|1.03037880E+00,3.21077480E+00,1.03982690E+01,7.59227210E-04|1.16235710E+00," & _
"3.45375690E+00,1.42424270E+01,9.05556190E-04|1.37518850E+00,4.04309520E+00,1.94001540E+01,1.16095850E-03|1.57480410E+00,4.71765200E+00,2.47951770E+01,1.54916600E-03|1.82959840E+00,5.92330780E+00,2.99356540E+01,2.31568240E-03|2.00189370E+00,6.59290820E+00,4.08562580E+01,3.05897210E-03|"

'-------------------------------------------------------------------------------------------------------------
ii = 1
'-------------------------------------------------------------------------------------------------------------
For i = 1 To NROWS
'-------------------------------------------------------------------------------------------------------------
    For j = 1 To NCOLUMNS - 1
        jj = InStr(ii, LINE_STR, DELIM_CHR)
        TEMP_STR = Mid(LINE_STR, ii, jj - ii)
        CNT_MATRIX(i, j) = CDec(TEMP_STR)
        ii = jj + Len(DELIM_CHR)
    Next j
    jj = InStr(ii, LINE_STR, TAB_CHR)
    TEMP_STR = Mid(LINE_STR, ii, jj - ii)
    CNT_MATRIX(i, NCOLUMNS) = CDec(TEMP_STR)
    ii = jj + Len(TAB_CHR)
'-------------------------------------------------------------------------------------------------------------
Next i
'-------------------------------------------------------------------------------------------------------------

READ_CNT_DAT_FUNC = CNT_MATRIX

Exit Function
ERROR_LABEL:
READ_CNT_DAT_FUNC = Err.number
End Function

Private Function READ_CT_DAT_FUNC() 'Perfect

Dim i As Long
Dim j As Long

Dim ii As Long
Dim jj As Long

Const DELIM_CHR As String = ","
Const TAB_CHR As String = "|"

Const NROWS As Long = 221
Const NCOLUMNS As Long = 5

Dim TEMP_STR As String
Dim LINE_STR As String


On Error GoTo ERROR_LABEL

ReDim CT_MATRIX(1 To NROWS, 1 To NCOLUMNS) As Double

LINE_STR = _
"-5.12922490E+00,-2.67194050E+01,-7.15050580E+01,-1.81322040E+03,3.38859020E-03|-4.97728430E+00,-2.30234960E+01,-1.06259100E+02,-8.24402180E+02,2.66363230E-03|-4.76750560E+00,-1.89373580E+01,-1.06179910E+02,-2.84376770E+02,1.77156830E-03|-4.59531350E+00,-1.67642090E+01,-6.83012280E+01,-4.62161380E+02,1.22677760E-03|-4.41751490E+00,-1.41249480E+01,-6.00384640E+01,-3.23384400E+02,9.37776910E-04|-4.30741250E+00,-1.29459020E+01,-4.46832630E+01,-3.36477430E+02,7.94683310E-04|-4.22757280E+00,-1.20306930E+01,-3.65660980E+01,-3.52386190E+02,6.91135800E-04|-4.16388430E+00,-1.13500530E+01,-3.27749430E+01,-2.92778500E+02,6.30127870E-04|-4.11118040E+00,-1.07381180E+01,-3.27935660E+01,-2.22429180E+02,5.84450690E-04|-4.06555160E+00,-1.03554850E+01,-2.49410820E+01,-2.69167190E+02,5.54628290E-04|-4.02571600E+00,-9.90516150E+00,-2.43429940E+01,-2.38231650E+02,5.33466110E-04|-3.99033820E+00,-9.49809730E+00,-2.44696110E+01,-2.02110840E+02,5.10618290E-04|-3.95797030E+00,-9.19695740E+00," & _
"-2.27027010E+01,-1.88391980E+02,4.89411120E-04|-3.83042970E+00,-7.99686310E+00,-1.53712820E+01,-1.86228060E+02,4.15781110E-04|-3.73597460E+00,-7.16831220E+00,-1.13429410E+01,-1.64673620E+02,3.72780250E-04|-3.66043480E+00,-6.49610920E+00,-9.80844110E+00,-1.38654350E+02,3.34838020E-04|-3.59714520E+00,-5.95847300E+00,-7.66980070E+00,-1.39807170E+02,3.17484850E-04|-3.54208860E+00,-5.53689530E+00,-4.92098220E+00,-1.51271420E+02,2.99988830E-04|-3.49338730E+00,-5.16180700E+00,-3.87343240E+00,-1.41616120E+02,2.82301150E-04|-3.44983120E+00,-4.79453740E+00,-4.51039580E+00,-1.18199840E+02,2.65933090E-04|-3.40982710E+00,-4.51196860E+00,-3.66303740E+00,-1.09969420E+02,2.60064140E-04|-3.37310390E+00,-4.25305400E+00,-2.68105390E+00,-1.11477160E+02,2.50159610E-04|-3.33892690E+00,-4.01614540E+00,-2.27938120E+00,-1.05992320E+02,2.43768950E-04|-3.30719580E+00,-3.78375820E+00,-2.68402880E+00,-8.85688170E+01,2.36314030E-04|-3.27726980E+00,-3.59896290E+00,-1.40969290E+00,-9.78865570E+01,2.31051110E-04|"

LINE_STR = LINE_STR & _
"-3.24906380E+00,-3.40048490E+00,-1.86471370E+00,-8.20412480E+01,2.23744660E-04|-3.22224960E+00,-3.22981640E+00,-1.64974710E+00,-7.67701460E+01,2.15994030E-04|-3.19667010E+00,-3.09145180E+00,-4.66486070E-01,-8.48321600E+01,2.15101580E-04|-3.17235410E+00,-2.94427370E+00,-2.51223620E-01,-7.98411180E+01,2.11963970E-04|-3.14904990E+00,-2.79846150E+00,-6.08411070E-01,-6.72482710E+01,2.03927520E-04|-3.12660080E+00,-2.68019930E+00,1.80644680E-01,-7.05668000E+01,1.98677220E-04|-3.10506140E+00,-2.55561460E+00,3.04830300E-01,-6.63313560E+01,1.96575210E-04|-3.08426310E+00,-2.43255360E+00,5.52377510E-02,-5.74413120E+01,1.93633020E-04|-3.06412400E+00,-2.35160580E+00,1.93699310E+00,-7.84144100E+01,1.89386480E-04|-3.04472840E+00,-2.23877490E+00,1.94737490E+00,-7.65245630E+01,1.89111440E-04|-3.02585310E+00,-2.13803200E+00,2.08956660E+00,-7.55362840E+01,1.87752060E-04|-3.00757430E+00,-2.03466680E+00,1.90877710E+00,-6.98415880E+01,1.83628160E-04|-2.98977210E+00,-1.94068360E+00,1.67261320E+00," & _
"-6.11914330E+01,1.81880850E-04|-2.97241860E+00,-1.86300550E+00,2.19824390E+00,-6.29914220E+01,1.79531840E-04|-2.95561520E+00,-1.77641910E+00,2.39430490E+00,-6.33232570E+01,1.77399960E-04|-2.93912700E+00,-1.70198140E+00,2.90113930E+00,-6.74014460E+01,1.76316210E-04|-2.92309150E+00,-1.61944640E+00,2.91000920E+00,-6.57349670E+01,1.73833080E-04|-2.90735420E+00,-1.55444660E+00,3.37149720E+00,-6.84479220E+01,1.71323970E-04|-2.89204210E+00,-1.47461170E+00,3.28729200E+00,-6.53477570E+01,1.70653860E-04|-2.87700720E+00,-1.39996560E+00,3.17623130E+00,-6.23575620E+01,1.68545060E-04|-2.86227670E+00,-1.32486690E+00,2.92483750E+00,-5.71175550E+01,1.67758770E-04|-2.84781440E+00,-1.25760210E+00,3.02756820E+00,-5.74902930E+01,1.68107000E-04|-2.83363680E+00,-1.18327250E+00,2.57749570E+00,-5.04172680E+01,1.67170410E-04|-2.81966930E+00,-1.13516720E+00,3.28424250E+00,-5.68669360E+01,1.65075660E-04|-2.80592290E+00,-1.08978310E+00,4.09878620E+00,-6.58450430E+01,1.64099380E-04|"

LINE_STR = LINE_STR & _
"-2.79242440E+00,-1.03578430E+00,4.31391950E+00,-6.64519530E+01,1.61224740E-04|-2.77915940E+00,-9.81162760E-01,4.40461390E+00,-6.56211840E+01,1.60216260E-04|-2.76609180E+00,-9.35653300E-01,4.95134560E+00,-7.18030480E+01,1.60280870E-04|-2.75329500E+00,-8.76030230E-01,4.75957280E+00,-6.84317770E+01,1.60260790E-04|-2.74057700E+00,-8.39018380E-01,5.45786230E+00,-7.49616380E+01,1.59593720E-04|-2.72814030E+00,-7.78416200E-01,4.91749470E+00,-6.53320970E+01,1.58219150E-04|-2.71584860E+00,-7.27391690E-01,4.90948460E+00,-6.41845970E+01,1.56751190E-04|-2.70370320E+00,-6.72686680E-01,4.49446290E+00,-5.65743530E+01,1.55964010E-04|-2.69167300E+00,-6.29735280E-01,4.59154870E+00,-5.56342190E+01,1.55963070E-04|-2.67986530E+00,-5.76946900E-01,4.26752440E+00,-5.00388140E+01,1.55746820E-04|-2.66814780E+00,-5.36575850E-01,4.43897380E+00,-5.06523470E+01,1.54280070E-04|-2.65663140E+00,-4.84718370E-01,4.05750160E+00,-4.42613540E+01,1.55278840E-04|-2.64521720E+00,-4.40068390E-01,4.08399190E+00,-4.41016830E+01," & _
"1.55186210E-04|-2.63392880E+00,-3.90808400E-01,3.59229500E+00,-3.54871370E+01,1.53036400E-04|-2.62269200E+00,-3.56808000E-01,3.96099520E+00,-3.96280370E+01,1.53247450E-04|-2.61162680E+00,-3.15858960E-01,3.89180790E+00,-3.73373590E+01,1.54110090E-04|-2.60070080E+00,-2.75080440E-01,3.88250030E+00,-3.72387950E+01,1.52602440E-04|-2.58986040E+00,-2.35958150E-01,3.83747610E+00,-3.57030140E+01,1.52060380E-04|-2.57910240E+00,-1.92751040E-01,3.61806770E+00,-3.28544710E+01,1.52333460E-04|-2.56847800E+00,-1.44224280E-01,2.94974930E+00,-2.32651650E+01,1.51873710E-04|-2.55789070E+00,-1.17081150E-01,3.48658370E+00,-3.02276540E+01,1.51466810E-04|-2.54744790E+00,-7.89283350E-02,3.37873420E+00,-2.85119270E+01,1.50840940E-04|-2.53705100E+00,-4.60020030E-02,3.49112930E+00,-2.98083410E+01,1.51350050E-04|-2.52673460E+00,-8.84333680E-03,3.25820710E+00,-2.55738310E+01,1.51467180E-04|-2.51651010E+00,3.02396660E-02,2.94001200E+00,-2.13693290E+01,1.50711370E-04|"

LINE_STR = LINE_STR & _
"-2.50634860E+00,5.71891110E-02,3.18155060E+00,-2.34567970E+01,1.50691880E-04|-2.49631730E+00,9.66844670E-02,2.83680080E+00,-1.88104000E+01,1.48968260E-04|-2.48635520E+00,1.34426780E-01,2.55918940E+00,-1.51522880E+01,1.48035790E-04|-2.47641340E+00,1.67070050E-01,2.39113640E+00,-1.22134980E+01,1.46420680E-04|-2.46651450E+00,1.98300710E-01,2.29718350E+00,-9.95918310E+00,1.46319120E-04|-2.45670760E+00,2.26018280E-01,2.42658590E+00,-1.15597080E+01,1.45514350E-04|-2.44692980E+00,2.50586040E-01,2.66706230E+00,-1.44238300E+01,1.45008810E-04|-2.43722320E+00,2.78010610E-01,2.67438910E+00,-1.34397080E+01,1.44679200E-04|-2.42762240E+00,3.13954620E-01,2.26215070E+00,-7.51220920E+00,1.44340480E-04|-2.41802300E+00,3.38469500E-01,2.37236210E+00,-8.15120600E+00,1.44681250E-04|-2.40850100E+00,3.71960420E-01,2.10089090E+00,-5.01659560E+00,1.43648860E-04|-2.39900180E+00,4.01270840E-01,1.85032100E+00,9.46709590E-02,1.44715420E-04|-2.38955190E+00,4.23571060E-01,2.05640430E+00,-2.09265510E+00,1.43841850E-04|" & _
"-2.38011780E+00,4.44831410E-01,2.31720420E+00,-5.75028420E+00,1.43493280E-04|-2.37073860E+00,4.67992630E-01,2.45109740E+00,-7.86906640E+00,1.42891310E-04|-2.36139450E+00,4.84531940E-01,2.84962250E+00,-1.24306690E+01,1.42461230E-04|-2.35213150E+00,5.12248190E-01,2.71723240E+00,-1.08058970E+01,1.43119510E-04|-2.34286150E+00,5.34990210E-01,2.73932950E+00,-1.04771700E+01,1.43738530E-04|-2.33365940E+00,5.66586890E-01,2.35618880E+00,-5.66974050E+00,1.43725500E-04|-2.32445500E+00,5.85372930E-01,2.57314500E+00,-8.37672620E+00,1.42752710E-04|-2.31529390E+00,6.10500060E-01,2.35447780E+00,-4.50723840E+00,1.42215040E-04|-2.30615450E+00,6.33214020E-01,2.24907960E+00,-2.31682240E+00,1.41502620E-04|-2.29707670E+00,6.57126800E-01,2.16283460E+00,-1.10152450E+00,1.41926150E-04|-2.28798660E+00,6.78729820E-01,2.05705380E+00,1.12348090E+00,1.41122280E-04|-2.27890510E+00,6.89169110E-01,2.69274330E+00,-7.40468930E+00,1.40247010E-04|"

LINE_STR = LINE_STR & _
"-2.26984880E+00,7.08284410E-01,2.74441580E+00,-7.47371310E+00,1.39963830E-04|-2.26083680E+00,7.33358320E-01,2.46635810E+00,-3.23732650E+00,1.40660220E-04|-2.25185330E+00,7.52219490E-01,2.57539420E+00,-4.66439760E+00,1.40967560E-04|-2.24287680E+00,7.70802240E-01,2.71532510E+00,-6.65071020E+00,1.39651350E-04|-2.23394500E+00,7.93823990E-01,2.60317050E+00,-5.24268950E+00,1.38419710E-04|-2.22500640E+00,8.16299510E-01,2.48286760E+00,-3.84139470E+00,1.37866480E-04|-2.21608110E+00,8.40617530E-01,2.21398880E+00,-1.66628170E-02,1.37013530E-04|-2.20715780E+00,8.56849050E-01,2.40740560E+00,-2.46243600E+00,1.38049430E-04|-2.19826000E+00,8.76492370E-01,2.42633140E+00,-3.05695450E+00,1.37978740E-04|-2.18936930E+00,8.97171470E-01,2.30493540E+00,-9.17618640E-01,1.38184530E-04|-2.18047230E+00,9.14073900E-01,2.39267280E+00,-2.17441090E+00,1.38406770E-04|-2.17156330E+00,9.33496570E-01,2.30610150E+00,-9.05841520E-01,1.37615970E-04|-2.16262760E+00,9.39719630E-01,2.83456740E+00,-7.16593280E+00,1.38323640E-04|" & _
"-2.15375430E+00,9.57887910E-01,2.92072760E+00,-8.90256720E+00,1.38073260E-04|-2.14487250E+00,9.76851270E-01,2.82260550E+00,-7.33109950E+00,1.37479150E-04|-2.13602340E+00,1.00180440E+00,2.50812610E+00,-3.78346330E+00,1.36080010E-04|-2.12714100E+00,1.02213810E+00,2.35916070E+00,-1.75527500E+00,1.36472180E-04|-2.11825720E+00,1.03769070E+00,2.45503520E+00,-2.80451440E+00,1.36066370E-04|-2.10935390E+00,1.05455030E+00,2.48325970E+00,-3.64853190E+00,1.36639960E-04|-2.10043920E+00,1.07075210E+00,2.51696200E+00,-4.38212380E+00,1.37923040E-04|-2.09151770E+00,1.07662100E+00,3.11975240E+00,-1.28230560E+01,1.38443320E-04|-2.08255770E+00,1.08934510E+00,3.18808120E+00,-1.31585460E+01,1.38638150E-04|-2.07363420E+00,1.10941160E+00,2.96821000E+00,-1.04691900E+01,1.37876020E-04|-2.06466550E+00,1.11888820E+00,3.38810700E+00,-1.64437750E+01,1.39297720E-04|-2.05569780E+00,1.13543940E+00,3.30646860E+00,-1.50439460E+01,1.38964610E-04|"

LINE_STR = LINE_STR & _
"-2.04671860E+00,1.15404500E+00,3.14192420E+00,-1.26758500E+01,1.40194310E-04|-2.03771190E+00,1.17015410E+00,3.13458580E+00,-1.28827300E+01,1.40369600E-04|-2.02868650E+00,1.18404890E+00,3.24059070E+00,-1.42773270E+01,1.41008530E-04|-2.01960660E+00,1.19082540E+00,3.69978720E+00,-2.07963410E+01,1.39972800E-04|-2.01050350E+00,1.19868520E+00,4.09488730E+00,-2.64149700E+01,1.40013600E-04|-2.00141010E+00,1.21748240E+00,3.83871600E+00,-2.28565740E+01,1.39146000E-04|-1.99226080E+00,1.22584900E+00,4.14157150E+00,-2.65620920E+01,1.38265480E-04|-1.98310550E+00,1.24212600E+00,4.12815580E+00,-2.66619640E+01,1.38679070E-04|-1.97393810E+00,1.26385680E+00,3.79814070E+00,-2.26139340E+01,1.38274430E-04|-1.96468170E+00,1.26621240E+00,4.44683850E+00,-3.14965350E+01,1.38177970E-04|-1.95540130E+00,1.28188930E+00,4.28578280E+00,-2.89486900E+01,1.36914890E-04|-1.94609050E+00,1.29496510E+00,4.33249660E+00,-2.97529160E+01,1.36956020E-04|-1.93673250E+00,1.30594000E+00,4.51481270E+00,-3.21561380E+01,1.36536500E-04|" & _
"-1.92734490E+00,1.31964730E+00,4.49705270E+00,-3.21038020E+01,1.36855570E-04|-1.91794200E+00,1.33901940E+00,4.23898640E+00,-2.84742530E+01,1.36838700E-04|-1.90849740E+00,1.36168520E+00,3.80769950E+00,-2.30244840E+01,1.36721940E-04|-1.89896330E+00,1.37924490E+00,3.58560950E+00,-1.99669220E+01,1.37532570E-04|-1.88939090E+00,1.40271740E+00,3.02189520E+00,-1.27272310E+01,1.38280680E-04|-1.87971540E+00,1.41144570E+00,3.22824610E+00,-1.49331040E+01,1.39223970E-04|-1.87001790E+00,1.42686650E+00,3.14720970E+00,-1.42325560E+01,1.40064290E-04|-1.86023680E+00,1.43560020E+00,3.45934360E+00,-1.87907990E+01,1.40160040E-04|-1.85042580E+00,1.45008480E+00,3.38637640E+00,-1.74347970E+01,1.40592730E-04|-1.84052550E+00,1.46623670E+00,3.25400390E+00,-1.61315790E+01,1.41283700E-04|-1.83054380E+00,1.48075860E+00,3.19376630E+00,-1.54619950E+01,1.42141320E-04|-1.82049910E+00,1.49481780E+00,3.24972260E+00,-1.69272320E+01,1.40758050E-04|"

LINE_STR = LINE_STR & _
"-1.81035100E+00,1.50977920E+00,3.19885940E+00,-1.64256510E+01,1.41394160E-04|-1.80012440E+00,1.52500880E+00,3.06540100E+00,-1.47283350E+01,1.40273160E-04|-1.78978800E+00,1.53086190E+00,3.40853400E+00,-1.86676210E+01,1.40371220E-04|-1.77937540E+00,1.53977140E+00,3.60012420E+00,-2.05865800E+01,1.40817750E-04|-1.76886690E+00,1.55150740E+00,3.62505240E+00,-2.01232360E+01,1.41785410E-04|-1.75828360E+00,1.56579760E+00,3.71085700E+00,-2.15303770E+01,1.41843050E-04|-1.74762150E+00,1.58889130E+00,3.37095760E+00,-1.81280330E+01,1.42751690E-04|-1.73681020E+00,1.60753220E+00,3.06282660E+00,-1.36981750E+01,1.44413640E-04|-1.72588020E+00,1.62171580E+00,2.94654030E+00,-1.08315160E+01,1.44147460E-04|-1.71477890E+00,1.62630070E+00,3.33479020E+00,-1.45676240E+01,1.44876480E-04|-1.70360550E+00,1.63850910E+00,3.59645220E+00,-1.94786590E+01,1.45809100E-04|-1.69230350E+00,1.65750110E+00,3.38754500E+00,-1.64710740E+01,1.46752540E-04|-1.68081780E+00,1.67350870E+00,3.34701120E+00,-1.59790850E+01,1.47796630E-04|" & _
"-1.66922080E+00,1.69763180E+00,2.91912070E+00,-1.11056850E+01,1.48459190E-04|-1.65744990E+00,1.71267200E+00,3.01903910E+00,-1.25443510E+01,1.49119570E-04|-1.64552030E+00,1.73211960E+00,2.78435570E+00,-8.89518040E+00,1.51331370E-04|-1.63342020E+00,1.75057710E+00,2.71408150E+00,-8.41346830E+00,1.51142200E-04|-1.62111750E+00,1.76472630E+00,2.83015500E+00,-9.39617230E+00,1.53022090E-04|-1.60863110E+00,1.78082170E+00,2.86139780E+00,-9.35381990E+00,1.53945670E-04|-1.59590470E+00,1.79668120E+00,2.98382010E+00,-1.11038410E+01,1.55187150E-04|-1.58300720E+00,1.81928770E+00,2.65039630E+00,-5.56307000E+00,1.55004750E-04|-1.56983720E+00,1.83411560E+00,2.73959500E+00,-5.13855620E+00,1.56876230E-04|-1.55642820E+00,1.85621270E+00,2.50121400E+00,-7.87160520E-01,1.58828360E-04|-1.54276070E+00,1.87737140E+00,2.34358650E+00,2.84724360E+00,1.60393720E-04|-1.52880020E+00,1.89566520E+00,2.29803920E+00,5.65953260E+00,1.62328410E-04|"

LINE_STR = LINE_STR & _
"-1.51453610E+00,1.90937330E+00,2.64144260E+00,2.35178650E+00,1.63653450E-04|-1.49996570E+00,1.92606940E+00,2.80475120E+00,1.68944300E+00,1.65661390E-04|-1.48508390E+00,1.95274550E+00,2.51698650E+00,7.64608260E+00,1.66415690E-04|-1.46984490E+00,1.97508830E+00,2.50259280E+00,9.88370330E+00,1.70349110E-04|-1.45415160E+00,1.98475680E+00,3.15048060E+00,4.27531720E+00,1.71417820E-04|-1.43810400E+00,2.00991170E+00,3.21571480E+00,4.83281790E+00,1.74480190E-04|-1.42156280E+00,2.02399760E+00,3.80390140E+00,-2.81533040E-01,1.76562340E-04|-1.40461280E+00,2.05434910E+00,3.55632720E+00,5.77513710E+00,1.76994200E-04|-1.38714000E+00,2.09165550E+00,3.07329650E+00,1.47675520E+01,1.78920360E-04|-1.36907550E+00,2.11139110E+00,3.48661910E+00,1.24626130E+01,1.82685880E-04|-1.35034170E+00,2.12364480E+00,4.39464260E+00,2.42001960E+00,1.86010540E-04|-1.33107190E+00,2.16296720E+00,4.00262670E+00,9.93162850E+00,1.87214000E-04|-1.31114320E+00,2.19948850E+00,4.07906450E+00,1.02407560E+01,1.90610320E-04|" & _
"-1.29040210E+00,2.24204360E+00,3.77917770E+00,1.54854020E+01,1.93886470E-04|-1.26890990E+00,2.30085850E+00,3.00057400E+00,2.54151000E+01,1.99111090E-04|-1.24640240E+00,2.32862000E+00,3.77095290E+00,1.70589530E+01,2.03956170E-04|-1.22288620E+00,2.36342950E+00,4.27750970E+00,1.12508070E+01,2.09446300E-04|-1.19824250E+00,2.40900330E+00,4.39322760E+00,8.90325030E+00,2.14056040E-04|-1.17231180E+00,2.44588870E+00,4.75689500E+00,6.81775650E+00,2.20299620E-04|-1.14501720E+00,2.49129670E+00,5.01526980E+00,2.94403210E+00,2.25634750E-04|-1.11619530E+00,2.54195240E+00,5.03101480E+00,2.60899190E+00,2.33226390E-04|-1.08569920E+00,2.60047770E+00,4.69510570E+00,7.81801450E+00,2.43286140E-04|-1.05323750E+00,2.68449600E+00,3.17434900E+00,2.49841980E+01,2.51898640E-04|-1.01850680E+00,2.76049400E+00,2.23214650E+00,3.49955790E+01,2.61082190E-04|-9.81059830E-01,2.82052070E+00,1.84240270E+00,4.24934380E+01,2.71920170E-04|"

LINE_STR = LINE_STR & _
"-9.40301120E-01,2.86365850E+00,2.36529300E+00,3.42680500E+01,2.81741680E-04|-8.96119410E-01,2.94104060E+00,1.59845430E+00,4.44460990E+01,2.92447180E-04|-8.47345160E-01,3.02849740E+00,3.66551170E-01,6.08662800E+01,3.03645790E-04|-7.93016870E-01,3.10324740E+00,-7.03830020E-02,6.79197850E+01,3.22595900E-04|-7.31346430E-01,3.16997690E+00,5.72387050E-01,5.64553040E+01,3.43599110E-04|-6.59882220E-01,3.23975330E+00,1.00386030E+00,5.55906830E+01,3.62815850E-04|-5.74989080E-01,3.37045490E+00,-8.68534770E-01,9.01628690E+01,4.03710430E-04|-4.68826160E-01,3.45656920E+00,1.08569590E+00,7.39448090E+01,4.58393750E-04|-3.25425030E-01,3.64673860E+00,9.72780760E-01,8.00960550E+01,5.35594210E-04|-2.89270490E-01,3.66679540E+00,2.67984440E+00,5.84986170E+01,5.46541760E-04|-2.49361610E-01,3.74255530E+00,1.90147210E+00,6.95175020E+01,5.65650840E-04|-2.04498540E-01,3.81835210E+00,1.21149510E+00,8.14932440E+01,5.90046850E-04|-1.53898250E-01,3.86862850E+00,3.77043850E+00,4.81326140E+01,6.30122890E-04|" & _
"-9.46164380E-02,3.90075530E+00,8.24102870E+00,-1.74593600E+01,6.61431450E-04|-2.38625490E-02,4.02078080E+00,7.14361520E+00,2.05700700E+01,7.50750380E-04|6.55207180E-02,4.15949450E+00,6.74793750E+00,4.63216450E+01,8.27198580E-04|1.88340400E-01,4.34597880E+00,9.74470540E+00,8.08549810E+00,9.62688520E-04|3.88631510E-01,4.65313640E+00,1.49688410E+01,-8.35534530E+00,1.32634630E-03|5.75870250E-01,5.07059400E+00,2.17325950E+01,-6.30422150E+01,1.74999560E-03|8.19045470E-01,5.25594690E+00,5.60078460E+01,-4.11789440E+02,2.62992910E-03|9.77209560E-01,6.61049170E+00,3.24251530E+01,-1.57352290E+02,3.42968610E-03|"
'-------------------------------------------------------------------------------------------------------------
ii = 1
'-------------------------------------------------------------------------------------------------------------
For i = 1 To NROWS
'-------------------------------------------------------------------------------------------------------------
    For j = 1 To NCOLUMNS - 1
        jj = InStr(ii, LINE_STR, DELIM_CHR)
        TEMP_STR = Mid(LINE_STR, ii, jj - ii)
        CT_MATRIX(i, j) = CDec(TEMP_STR)
        ii = jj + Len(DELIM_CHR)
    Next j
    jj = InStr(ii, LINE_STR, TAB_CHR)
    TEMP_STR = Mid(LINE_STR, ii, jj - ii)
    CT_MATRIX(i, NCOLUMNS) = CDec(TEMP_STR)
    ii = jj + Len(TAB_CHR)
'-------------------------------------------------------------------------------------------------------------
Next i
'-------------------------------------------------------------------------------------------------------------

READ_CT_DAT_FUNC = CT_MATRIX

Exit Function
ERROR_LABEL:
READ_CT_DAT_FUNC = Err.number
End Function


Private Function READ_NC_DAT_FUNC() 'Perfect

Dim i As Long
Dim j As Long

Dim ii As Long
Dim jj As Long

Const DELIM_CHR As String = ","
Const TAB_CHR As String = "|"

Const NROWS As Long = 221
Const NCOLUMNS As Long = 4

Dim TEMP_STR As String
Dim LINE_STR As String


On Error GoTo ERROR_LABEL

ReDim NC_MATRIX(1 To NROWS, 1 To NCOLUMNS) As Double

LINE_STR = _
"-3.89296810E+00,-1.07114370E+01,-5.56199550E+01,3.31881920E-03|-3.71858720E+00,-9.58969210E+00,-3.41665320E+01,2.49514890E-03|-3.47591640E+00,-7.87460510E+00,-1.74906570E+01,1.55591560E-03|-3.28427080E+00,-6.41089120E+00,-1.19837550E+01,1.14459640E-03|-3.08304570E+00,-5.05676330E+00,-6.88471430E+00,8.41007650E-04|-2.95986020E+00,-4.25847160E+00,-5.81424630E+00,7.10499340E-04|-2.86988250E+00,-3.72895500E+00,-5.87149710E+00,6.23911510E-04|-2.79770770E+00,-3.41464270E+00,-3.76153000E+00,5.86932440E-04|-2.73800500E+00,-3.12084980E+00,-3.04628830E+00,5.25764070E-04|-2.68684360E+00,-2.87224760E+00,-2.54385630E+00,4.90529760E-04|-2.64168520E+00,-2.67934020E+00,-1.87697830E+00,4.68043160E-04|-2.60125970E+00,-2.50000040E+00,-1.63079330E+00,4.43461340E-04|-2.56494460E+00,-2.31273990E+00,-2.03066050E+00,4.30390030E-04|-2.42061780E+00,-1.72918860E+00,-1.34664930E+00,3.64674750E-04|-2.31333040E+00,-1.34236150E+00,-1.34322570E+00,3.21071640E-04|-2.22731430E+00,-1.07462640E+00," & _
"-8.76242700E-01,2.97897550E-04|-2.15490470E+00,-8.46576280E-01,-1.07801880E+00,2.78305830E-04|-2.09216550E+00,-6.77443160E-01,-1.04306970E+00,2.60276020E-04|-2.03645840E+00,-5.41924870E-01,-9.69651940E-01,2.52491340E-04|-1.98641460E+00,-4.16168900E-01,-1.13340450E+00,2.40799950E-04|-1.94076840E+00,-3.30252310E-01,-6.78214260E-01,2.33230220E-04|-1.89885850E+00,-2.29707730E-01,-8.56408490E-01,2.20751410E-04|-1.85984880E+00,-1.46397280E-01,-1.04288060E+00,2.13349550E-04|-1.82335730E+00,-8.96638030E-02,-8.08559450E-01,2.09630100E-04|-1.78913980E+00,-3.64668650E-02,-6.32438960E-01,2.08342320E-04|"

LINE_STR = LINE_STR & _
"-1.75679930E+00,1.14026540E-02,-5.10664970E-01,2.05399330E-04|-1.72621050E+00,6.62209560E-02,-6.37080030E-01,1.99647620E-04|-1.69698380E+00,1.02368240E-01,-4.66727870E-01,1.95741320E-04|-1.66908590E+00,1.42653320E-01,-4.50614180E-01,1.88880620E-04|-1.64242710E+00,1.87018230E-01,-6.60843350E-01,1.84265550E-04|-1.61672790E+00,2.16267890E-01,-6.46169910E-01,1.81280890E-04|-1.59205200E+00,2.47789880E-01,-6.35975020E-01,1.79549020E-04|-1.56821140E+00,2.73235620E-01,-6.16181570E-01,1.77734950E-04|-1.54527400E+00,2.99910060E-01,-6.05445650E-01,1.71421680E-04|-1.52306430E+00,3.31483330E-01,-7.65806370E-01,1.67692020E-04|-1.50148650E+00,3.51363870E-01,-6.90246910E-01,1.68871470E-04|-1.48057000E+00,3.74786670E-01,-7.29303880E-01,1.68722180E-04|-1.46025220E+00,3.94206570E-01,-7.04436010E-01,1.64017310E-04|-1.44042420E+00,4.10286810E-01,-6.43385180E-01,1.62246070E-04|-1.42107850E+00,4.21192000E-01,-5.55484000E-01,1.61931960E-04|-1.40224060E+00,4.37477460E-01,-5.76708200E-01," & _
"1.60978490E-04|-1.38378760E+00,4.45752980E-01,-4.32362790E-01,1.58038320E-04|-1.36577170E+00,4.55288830E-01,-3.51655690E-01,1.56949080E-04|-1.34810320E+00,4.63061810E-01,-3.10931840E-01,1.56035280E-04|-1.33091440E+00,4.79724730E-01,-3.91212600E-01,1.55570330E-04|-1.31395470E+00,4.88141750E-01,-3.80067230E-01,1.52367270E-04|-1.29739310E+00,4.98594370E-01,-3.54566980E-01,1.49095930E-04|-1.28110300E+00,5.04206310E-01,-2.55831700E-01,1.47934700E-04|-1.26512620E+00,5.14297880E-01,-2.88935570E-01,1.46555930E-04|-1.24940040E+00,5.21351230E-01,-2.78823090E-01,1.44126020E-04|"

LINE_STR = LINE_STR & _
"-1.23395220E+00,5.27565950E-01,-2.32175060E-01,1.44378380E-04|-1.21871680E+00,5.34809580E-01,-2.64062680E-01,1.44324080E-04|-1.20377020E+00,5.43385640E-01,-2.97290240E-01,1.42629550E-04|-1.18901550E+00,5.48172660E-01,-2.49007120E-01,1.41114210E-04|-1.17446770E+00,5.52639330E-01,-2.21290350E-01,1.40676480E-04|-1.16014440E+00,5.59195650E-01,-2.36231510E-01,1.39736850E-04|-1.14601770E+00,5.65176980E-01,-2.45654470E-01,1.39263840E-04|-1.13200960E+00,5.66032150E-01,-1.75242260E-01,1.39697800E-04|-1.11822920E+00,5.70110950E-01,-1.94398760E-01,1.39163430E-04|-1.10458660E+00,5.72097360E-01,-1.60749070E-01,1.38181670E-04|-1.09110740E+00,5.72348500E-01,-1.05242580E-01,1.37220180E-04|-1.07775560E+00,5.72296760E-01,-3.25507320E-02,1.36192880E-04|-1.06456840E+00,5.73023960E-01,1.43797490E-02,1.35534100E-04|-1.05152440E+00,5.73625900E-01,8.46391800E-02,1.35265270E-04|-1.03863170E+00,5.77259980E-01,7.35906400E-02,1.35191960E-04|-1.02584920E+00,5.80443400E-01,5.10600380E-02," & _
"1.33793160E-04|-1.01317000E+00,5.83859330E-01,6.12450250E-03,1.34283600E-04|-1.00062700E+00,5.89595350E-01,-6.49754460E-02,1.34032530E-04|-9.88163400E-01,5.89629040E-01,-3.89691700E-02,1.33886450E-04|-9.75808920E-01,5.92285430E-01,-9.20286080E-02,1.32562830E-04|-9.63539880E-01,5.90231180E-01,-2.52455370E-02,1.33501080E-04|-9.51354370E-01,5.92350810E-01,-6.69276980E-02,1.31666730E-04|-9.39271270E-01,5.94955230E-01,-1.17725640E-01,1.31098740E-04|-9.27257210E-01,5.94652510E-01,-9.88421510E-02,1.30758470E-04|-9.15382940E-01,5.98629480E-01,-1.33738700E-01,1.29919970E-04|"

LINE_STR = LINE_STR & _
"-9.03515910E-01,5.97360710E-01,-1.15752970E-01,1.30046510E-04|-8.91717400E-01,5.98275580E-01,-1.13763910E-01,1.29817790E-04|-8.79993580E-01,5.98546440E-01,-1.12629760E-01,1.29379300E-04|-8.68363930E-01,6.01259230E-01,-1.57005400E-01,1.28094820E-04|-8.56766300E-01,6.02352430E-01,-1.53411540E-01,1.27274900E-04|-8.45218210E-01,6.07247470E-01,-2.81427050E-01,1.28548990E-04|-8.33684460E-01,6.04147640E-01,-2.27594170E-01,1.28865270E-04|-8.22177540E-01,5.98824700E-01,-1.03693800E-01,1.27291100E-04|-8.10774770E-01,6.02667540E-01,-1.92369200E-01,1.27447910E-04|-7.99380130E-01,6.02036220E-01,-1.54556600E-01,1.28517570E-04|-7.87997800E-01,6.03109630E-01,-2.12365070E-01,1.28472800E-04|-7.76665560E-01,6.05121410E-01,-2.41959490E-01,1.28677830E-04|-7.65311520E-01,6.01610400E-01,-1.84359510E-01,1.27646270E-04|-7.54005090E-01,6.00056440E-01,-1.57587110E-01,1.27191620E-04|-7.42725500E-01,6.02630010E-01,-1.90737410E-01,1.27377330E-04|-7.31409500E-01,6.01697550E-01,-1.68099400E-01," & _
"1.29011650E-04|-7.20108210E-01,6.00738990E-01,-1.71185170E-01,1.29899010E-04|-7.08814380E-01,6.02813410E-01,-1.97346290E-01,1.30967560E-04|-6.97522540E-01,6.05741000E-01,-2.39005700E-01,1.32000270E-04|-6.86212420E-01,6.07090030E-01,-2.50078550E-01,1.32057420E-04|-6.74868990E-01,6.05951860E-01,-2.17803550E-01,1.33081730E-04|-6.63553100E-01,6.08799910E-01,-2.39076140E-01,1.33347340E-04|-6.52179060E-01,6.09344900E-01,-2.24027390E-01,1.33131130E-04|-6.40784700E-01,6.10450360E-01,-2.25256970E-01,1.34234810E-04|-6.29356680E-01,6.11654550E-01,-1.85557330E-01,1.35611640E-04|"

LINE_STR = LINE_STR & _
"-6.17849920E-01,6.09062240E-01,-1.02171600E-01,1.35243950E-04|-6.06340660E-01,6.11336450E-01,-9.35633120E-02,1.35740280E-04|-5.94799080E-01,6.12430940E-01,-1.80092530E-02,1.36779740E-04|-5.83171310E-01,6.13834240E-01,3.39766780E-02,1.38516270E-04|-5.71477640E-01,6.13839490E-01,1.29709470E-01,1.39192800E-04|-5.59746640E-01,6.18152020E-01,1.35674480E-01,1.38657350E-04|-5.47964150E-01,6.24250070E-01,1.26810840E-01,1.40277420E-04|-5.36057810E-01,6.27886520E-01,1.39226600E-01,1.41603620E-04|-5.24078230E-01,6.30030470E-01,2.00127400E-01,1.42638680E-04|-5.12030770E-01,6.34568090E-01,2.28498520E-01,1.44764710E-04|-4.99934820E-01,6.41029400E-01,2.36512260E-01,1.47455150E-04|-4.87746460E-01,6.48545690E-01,1.91839100E-01,1.48829840E-04|-4.75462580E-01,6.53912290E-01,2.06894620E-01,1.50630900E-04|-4.63125620E-01,6.61904500E-01,1.50848540E-01,1.50525650E-04|-4.50682360E-01,6.70324670E-01,1.00521720E-01,1.52613830E-04|-4.38165710E-01,6.79062780E-01,1.80431580E-02,1.51936560E-04|" & _
"-4.25530120E-01,6.85961660E-01,-3.43730750E-02,1.53242550E-04|-4.12789150E-01,6.87516270E-01,1.84992620E-03,1.52113740E-04|-3.99969010E-01,6.92625080E-01,-4.41060980E-02,1.52622410E-04|-3.87055870E-01,6.94701030E-01,-3.72235400E-02,1.52336480E-04|-3.74076970E-01,6.99151560E-01,-6.81046820E-02,1.53108490E-04|-3.61036000E-01,7.04185820E-01,-9.84763850E-02,1.54177300E-04|-3.47923190E-01,7.07306640E-01,-1.00851930E-01,1.55100720E-04|-3.34712670E-01,7.11040640E-01,-1.26712260E-01,1.55393960E-04|-3.21421070E-01,7.14785130E-01,-1.53185440E-01,1.56781980E-04|"

LINE_STR = LINE_STR & _
"-3.08031480E-01,7.13728910E-01,-1.09386380E-01,1.58867700E-04|-2.94630580E-01,7.20562820E-01,-2.28111980E-01,1.58691680E-04|-2.81090520E-01,7.21063030E-01,-2.08220060E-01,1.59869580E-04|-2.67543990E-01,7.30032350E-01,-3.75001150E-01,1.60052060E-04|-2.53847620E-01,7.30883400E-01,-3.64141540E-01,1.59824460E-04|-2.40099150E-01,7.33295740E-01,-3.97087470E-01,1.60642150E-04|-2.26274890E-01,7.35505000E-01,-3.88617380E-01,1.60282590E-04|-2.12364270E-01,7.35626850E-01,-3.77456950E-01,1.59069870E-04|-1.98350270E-01,7.36174670E-01,-3.74129040E-01,1.58813420E-04|-1.84289510E-01,7.36445060E-01,-3.43039830E-01,1.60104710E-04|-1.70119540E-01,7.36183750E-01,-2.98903530E-01,1.60914890E-04|-1.55868200E-01,7.33732460E-01,-2.27080900E-01,1.61223870E-04|-1.41521980E-01,7.32711650E-01,-2.17538430E-01,1.59859510E-04|-1.27100860E-01,7.30485680E-01,-1.69970320E-01,1.59984930E-04|-1.12587980E-01,7.27547930E-01,-1.00773340E-01,1.59986720E-04|-9.79760700E-02,7.28847780E-01,-1.08975800E-01," & _
"1.61405120E-04|-8.32872620E-02,7.30246380E-01,-1.01724130E-01,1.61104310E-04|-6.84994800E-02,7.34316620E-01,-1.85663600E-01,1.62827830E-04|-5.35963080E-02,7.33980930E-01,-1.56073470E-01,1.65727820E-04|-3.85919700E-02,7.33892180E-01,-1.75419390E-01,1.66223910E-04|-2.34703820E-02,7.32819120E-01,-1.87743820E-01,1.66979620E-04|-8.24140510E-03,7.32435520E-01,-1.66818160E-01,1.67359380E-04|7.09381880E-03,7.33116430E-01,-1.69261170E-01,1.70016410E-04|2.25585310E-02,7.33235360E-01,-1.69126390E-01,1.71344100E-04|3.81240400E-02,7.34236430E-01,-2.16316370E-01,1.71669190E-04|"

LINE_STR = LINE_STR & _
"5.38228760E-02,7.36956390E-01,-2.43742850E-01,1.71352160E-04|6.97239740E-02,7.34979010E-01,-2.12839840E-01,1.71788500E-04|8.57486480E-02,7.33530540E-01,-2.28876180E-01,1.73783090E-04|1.01947760E-01,7.27517240E-01,-1.40736150E-01,1.73679480E-04|1.18223740E-01,7.29199210E-01,-1.91814180E-01,1.74098130E-04|1.34661180E-01,7.28418730E-01,-2.05922640E-01,1.74040240E-04|1.51232020E-01,7.27482140E-01,-1.42825000E-01,1.74456530E-04|1.67997090E-01,7.22064660E-01,-1.59657750E-02,1.76989240E-04|1.84927700E-01,7.18740170E-01,6.60049760E-02,1.77781510E-04|2.02022740E-01,7.19343680E-01,2.56243590E-02,1.79555980E-04|2.19273040E-01,7.20392750E-01,2.55803800E-02,1.81845710E-04|2.36781020E-01,7.17487170E-01,9.40602520E-02,1.83646530E-04|2.54516360E-01,7.08858970E-01,2.50884850E-01,1.84731650E-04|2.72388550E-01,7.05036840E-01,3.74600850E-01,1.85850390E-04|2.90455300E-01,7.04691310E-01,4.36229790E-01,1.86690060E-04|3.08762890E-01,6.99913850E-01,5.70784470E-01,1.87318200E-04|3.27309860E-01," & _
"6.98176890E-01,6.53222550E-01,1.88814340E-04|3.46082010E-01,6.99544400E-01,6.52072220E-01,1.89328270E-04|3.65091600E-01,7.01183500E-01,6.67445630E-01,1.90364770E-04|3.84384500E-01,7.04795230E-01,6.38486920E-01,1.91048890E-04|4.03990540E-01,7.01275910E-01,7.73813190E-01,1.92546530E-04|4.23818590E-01,7.05785850E-01,7.76612710E-01,1.93579490E-04|4.44009810E-01,7.08451880E-01,7.65648430E-01,1.92708000E-04|4.64521230E-01,7.15046920E-01,6.69841530E-01,1.93996210E-04|4.85361960E-01,7.17272010E-01,7.29349900E-01,1.95574740E-04|"

LINE_STR = LINE_STR & _
"5.06500980E-01,7.31109110E-01,5.45675340E-01,1.98375280E-04|5.28117820E-01,7.28818880E-01,7.38903400E-01,1.99667280E-04|5.50064270E-01,7.35235770E-01,7.48364550E-01,1.99917040E-04|5.72448660E-01,7.41344380E-01,8.00910130E-01,1.99027370E-04|5.95310370E-01,7.43676250E-01,9.19786650E-01,2.02394380E-04|6.18695830E-01,7.46580140E-01,9.96431350E-01,2.05850530E-04|6.42497300E-01,7.63055790E-01,8.77710890E-01,2.05261290E-04|6.66917070E-01,7.69964440E-01,9.73620520E-01,2.08669440E-04|6.91831280E-01,7.85577660E-01,8.87540620E-01,2.09551340E-04|7.17449060E-01,8.00404330E-01,8.03080010E-01,2.12958250E-04|7.43772760E-01,8.12142800E-01,8.37231760E-01,2.13280530E-04|7.70808980E-01,8.25408210E-01,8.58957980E-01,2.18792410E-04|7.98657740E-01,8.35930590E-01,9.19652890E-01,2.24258730E-04|8.27451400E-01,8.43998370E-01,1.05155970E+00,2.27933820E-04|8.57060100E-01,8.60647300E-01,1.11678880E+00,2.31968730E-04|8.87736480E-01,8.76417500E-01,1.17963580E+00,2.39097000E-04|9.19527910E-01,8.99542670E-01," & _
"1.12307150E+00,2.39656760E-04|9.52597650E-01,9.15727360E-01,1.34460010E+00,2.42445770E-04|9.86858420E-01,9.46684990E-01,1.30832480E+00,2.43662510E-04|1.02267230E+00,9.72710070E-01,1.52542800E+00,2.48797610E-04|1.06027640E+00,1.00381430E+00,1.57918820E+00,2.53949770E-04|1.09981110E+00,1.03528140E+00,1.80385520E+00,2.57463260E-04|1.14157180E+00,1.06823340E+00,2.09454650E+00,2.59622310E-04|1.18574270E+00,1.10688150E+00,2.41427520E+00,2.62967800E-04|1.23298030E+00,1.14704140E+00,2.84972970E+00,2.66868700E-04|"

LINE_STR = LINE_STR & _
"1.28361010E+00,1.21111230E+00,2.94466120E+00,2.74748000E-04|1.33831680E+00,1.27556730E+00,3.38084550E+00,2.82729510E-04|1.39812510E+00,1.34926460E+00,3.80006200E+00,2.91856510E-04|1.46406420E+00,1.45448960E+00,4.11037140E+00,3.13854590E-04|1.53834550E+00,1.57136430E+00,4.61051000E+00,3.28108870E-04|1.62335690E+00,1.73144380E+00,5.17651440E+00,3.50164570E-04|1.72418040E+00,1.92698800E+00,6.35080300E+00,3.84378730E-04|1.84858340E+00,2.24589040E+00,7.23172070E+00,4.15617580E-04|2.01502480E+00,2.72647670E+00,8.84460310E+00,4.86017320E-04|2.05671360E+00,2.85273150E+00,9.61848470E+00,5.17968990E-04|2.10273470E+00,3.00081510E+00,1.03828660E+01,5.38981210E-04|2.15401260E+00,3.18226660E+00,1.10718900E+01,5.57721100E-04|2.21241080E+00,3.34544900E+00,1.28655210E+01,5.87844370E-04|2.27974200E+00,3.60150370E+00,1.40197400E+01,6.33545390E-04|2.36033550E+00,3.97075810E+00,1.45230690E+01,7.15691750E-04|2.46168520E+00,4.42581790E+00,1.55622430E+01,7.75966960E-04|2.60092170E+00,4.91971150E+00," & _
"2.26401370E+01,9.07870280E-04|2.82334970E+00,6.30453550E+00,2.64177220E+01,1.17165040E-03|3.03280320E+00,7.56319540E+00,3.50874910E+01,1.62400910E-03|3.29914050E+00,9.47489270E+00,5.33824130E+01,2.36434220E-03|3.47810860E+00,1.09601810E+01,6.57771050E+01,3.13826260E-03|"

'-------------------------------------------------------------------------------------------------------------
ii = 1
'-------------------------------------------------------------------------------------------------------------
For i = 1 To NROWS
'-------------------------------------------------------------------------------------------------------------
    For j = 1 To NCOLUMNS - 1
        jj = InStr(ii, LINE_STR, DELIM_CHR)
        TEMP_STR = Mid(LINE_STR, ii, jj - ii)
        NC_MATRIX(i, j) = CDec(TEMP_STR)
        ii = jj + Len(DELIM_CHR)
    Next j
    jj = InStr(ii, LINE_STR, TAB_CHR)
    TEMP_STR = Mid(LINE_STR, ii, jj - ii)
    NC_MATRIX(i, NCOLUMNS) = CDec(TEMP_STR)
    ii = jj + Len(TAB_CHR)
'-------------------------------------------------------------------------------------------------------------
Next i
'-------------------------------------------------------------------------------------------------------------

READ_NC_DAT_FUNC = NC_MATRIX

Exit Function
ERROR_LABEL:
READ_NC_DAT_FUNC = Err.number
End Function

Private Function READ_PROB_DAT_FUNC() 'Perfect

Dim i As Long

Dim ii As Long
Dim jj As Long

Const DELIM_CHR As String = ","
Const TAB_CHR As String = "|"

Const NROWS As Long = 221
Const NCOLUMNS As Long = 2

Dim TEMP_STR As String
Dim LINE_STR As String

On Error GoTo ERROR_LABEL

ReDim PROB_MATRIX(1 To NROWS, 1 To NCOLUMNS) As Double

LINE_STR = "0.0001,-3.71946953E+00|0.0002,-3.54018994E+00|0.0005,-3.29047907E+00|0.001,-3.09024472E+00|0.002,-2.87815055E+00|0.003,-2.74776539E+00|0.004,-2.65208655E+00|0.005,-2.57583451E+00|0.006,-2.51213351E+00|0.007,-2.45727279E+00|0.008,-2.40892405E+00|0.009,-2.36561391E+00|0.01,-2.32634193E+00|0.015,-2.17009074E+00|0.02,-2.05374818E+00|0.025,-1.95996108E+00|0.03,-1.88078957E+00|0.035,-1.81191353E+00|0.04,-1.75068635E+00|0.045,-1.69539817E+00|0.05,-1.64485300E+00|0.055,-1.59819137E+00|0.06,-1.55477210E+00|0.065,-1.51410404E+00|0.07,-1.47579158E+00|"
LINE_STR = LINE_STR & "0.075,-1.43953002E+00|0.08,-1.40507382E+00|0.085,-1.37220468E+00|0.09,-1.34075435E+00|0.095,-1.31057959E+00|0.1,-1.28155079E+00|0.105,-1.25356564E+00|0.11,-1.22652864E+00|0.115,-1.20036020E+00|0.12,-1.17498757E+00|0.125,-1.15034936E+00|0.13,-1.12639100E+00|0.135,-1.10306246E+00|0.14,-1.08032054E+00|0.145,-1.05812205E+00|0.15,-1.03643288E+00|0.155,-1.01522119E+00|0.16,-9.94457423E-01|0.165,-9.74114300E-01|0.17,-9.54164534E-01|0.175,-9.34589934E-01|0.18,-9.15365490E-01|0.185,-8.96473011E-01|0.19,-8.77896582E-01|0.195,-8.59618012E-01|"
LINE_STR = LINE_STR & "0.2,-8.41621386E-01|0.205,-8.23893060E-01|0.21,-8.06421667E-01|0.215,-7.89191290E-01|0.22,-7.72192834E-01|0.225,-7.55414931E-01|0.23,-7.38846211E-01|0.235,-7.22478717E-01|0.24,-7.06302217E-01|0.245,-6.90308752E-01|0.25,-6.74490366E-01|0.255,-6.58837962E-01|0.26,-6.43344720E-01|0.265,-6.28006092E-01|0.27,-6.12812983E-01|0.275,-5.97760845E-01|0.28,-5.82840585E-01|0.285,-5.68052201E-01|0.29,-5.53384325E-01|0.295,-5.38835820E-01|0.3,-5.24401003E-01|0.305,-5.10074187E-01|0.31,-4.95849690E-01|0.315,-4.81727511E-01|0.32,-4.67698555E-01|"
LINE_STR = LINE_STR & "0.325,-4.53762823E-01|0.33,-4.39913492E-01|0.335,-4.26148290E-01|0.34,-4.12462668E-01|0.345,-3.98855491E-01|0.35,-3.85321073E-01|0.355,-3.71856004E-01|0.36,-3.58459147E-01|0.365,-3.45125954E-01|0.37,-3.31854153E-01|0.375,-3.18639195E-01|0.38,-3.05481080E-01|0.385,-2.92375262E-01|0.39,-2.79319465E-01|0.395,-2.66311417E-01|0.4,-2.53346570E-01|0.405,-2.40426061E-01|0.41,-2.27545343E-01|0.415,-2.14702141E-01|0.42,-2.01894181E-01|0.425,-1.89118055E-01|0.43,-1.76373760E-01|0.435,-1.63659024E-01|0.44,-1.50969299E-01|0.445,-1.38304586E-01|"
LINE_STR = LINE_STR & "0.45,-1.25661472E-01|0.455,-1.13038823E-01|0.46,-1.00433226E-01|0.465,-8.78446826E-02|0.47,-7.52697815E-02|0.475,-6.27062491E-02|0.48,-5.01540853E-02|0.485,-3.76076059E-02|0.49,-2.50690846E-02|0.495,-1.25328370E-02|0.5,0.00000000E+00|0.505,1.25328370E-02|0.51,2.50690846E-02|0.515,3.76076059E-02|0.52,5.01540853E-02|0.525,6.27062491E-02|0.53,7.52697815E-02|0.535,8.78446826E-02|0.54,1.00433226E-01|0.545,1.13038823E-01|0.55,1.25661472E-01|0.555,1.38304586E-01|0.56,1.50969299E-01|0.565,1.63659024E-01|0.57,1.76373760E-01|"
LINE_STR = LINE_STR & "0.575,1.89118055E-01|0.58,2.01894181E-01|0.585,2.14702141E-01|0.59,2.27545343E-01|0.595,2.40426061E-01|0.6,2.53346570E-01|0.605,2.66311417E-01|0.61,2.79319465E-01|0.615,2.92375262E-01|0.62,3.05481080E-01|0.625,3.18639195E-01|0.63,3.31854153E-01|0.635,3.45125954E-01|0.64,3.58459147E-01|0.645,3.71856004E-01|0.65,3.85321073E-01|0.655,3.98855491E-01|0.66,4.12462668E-01|0.665,4.26148290E-01|0.67,4.39913492E-01|0.675,4.53762823E-01|0.68,4.67698555E-01|0.685,4.81727511E-01|0.69,4.95849690E-01|0.695,5.10074187E-01|"
LINE_STR = LINE_STR & "0.7,5.24401003E-01|0.705,5.38835820E-01|0.71,5.53384325E-01|0.715,5.68052201E-01|0.72,5.82840585E-01|0.725,5.97760845E-01|0.73,6.12812983E-01|0.735,6.28006092E-01|0.74,6.43344720E-01|0.745,6.58837962E-01|0.75,6.74490366E-01|0.755,6.90308752E-01|0.76,7.06302217E-01|0.765,7.22478717E-01|0.77,7.38846211E-01|0.775,7.55414931E-01|0.78,7.72192834E-01|0.785,7.89191290E-01|0.79,8.06421667E-01|0.795,8.23893060E-01|0.8,8.41621386E-01|0.805,8.59618012E-01|0.81,8.77896582E-01|0.815,8.96473011E-01|0.82,9.15365490E-01|"
LINE_STR = LINE_STR & "0.825,9.34589934E-01|0.83,9.54164534E-01|0.835,9.74114300E-01|0.84,9.94457423E-01|0.845,1.01522119E+00|0.85,1.03643288E+00|0.855,1.05812205E+00|0.86,1.08032054E+00|0.865,1.10306246E+00|0.87,1.12639100E+00|0.875,1.15034936E+00|0.88,1.17498757E+00|0.885,1.20036020E+00|0.89,1.22652864E+00|0.895,1.25356564E+00|0.9,1.28155079E+00|0.905,1.31057959E+00|0.91,1.34075435E+00|0.915,1.37220468E+00|0.92,1.40507382E+00|0.925,1.43953002E+00|0.93,1.47579158E+00|0.935,1.51410404E+00|0.94,1.55477210E+00|0.945,1.59819137E+00|"
LINE_STR = LINE_STR & "0.95,1.64485300E+00|0.955,1.69539817E+00|0.96,1.75068635E+00|0.965,1.81191353E+00|0.97,1.88078957E+00|0.975,1.95996108E+00|0.98,2.05374818E+00|0.985,2.17009074E+00|0.99,2.32634193E+00|0.991,2.36561391E+00|0.992,2.40892405E+00|0.993,2.45727279E+00|0.994,2.51213351E+00|0.995,2.57583451E+00|0.996,2.65208655E+00|0.997,2.74776539E+00|0.998,2.87815055E+00|0.999,3.09024472E+00|0.9995,3.29047907E+00|0.9998,3.54018994E+00|0.9999,3.71946953E+00|"

'-------------------------------------------------------------------------------------------------------------
ii = 1
'-------------------------------------------------------------------------------------------------------------
For i = 1 To NROWS
'-------------------------------------------------------------------------------------------------------------
    jj = InStr(ii, LINE_STR, DELIM_CHR)
    TEMP_STR = Mid(LINE_STR, ii, jj - ii)
    PROB_MATRIX(i, 1) = CDec(TEMP_STR)
    ii = jj + Len(DELIM_CHR)
    
    jj = InStr(ii, LINE_STR, TAB_CHR)
    TEMP_STR = Mid(LINE_STR, ii, jj - ii)
    PROB_MATRIX(i, 2) = CDec(TEMP_STR)
    ii = jj + Len(TAB_CHR)
'-------------------------------------------------------------------------------------------------------------
Next i
'-------------------------------------------------------------------------------------------------------------

READ_PROB_DAT_FUNC = PROB_MATRIX

Exit Function
ERROR_LABEL:
READ_PROB_DAT_FUNC = Err.number
End Function
