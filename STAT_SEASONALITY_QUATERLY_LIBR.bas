Attribute VB_Name = "STAT_SEASONALITY_QUATERLY_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
                    
Private PUB_XDATA_MATRIX As Variant
Private PUB_YDATA_VECTOR As Variant
Private Const PUB_EPSILON As Double = 2 ^ 52


'************************************************************************************
'************************************************************************************
'FUNCTION      : QUARTERLY_SEASONALITY_SIMULATION_FUNC
'DESCRIPTION   :

'This routine implements a seasonal variation and shows how regression can be used
'to deconstruct the data to a seasonally adjusted state. Unfortunately, seasonal
'adjustment relies heavily on knowing the seasonality process.

'LIBRARY       : STATISTICS
'GROUP         : SEASONALITY
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function QUARTERLY_SEASONALITY_SIMULATION_FUNC( _
Optional ByVal STARTING_PERIOD As Long = 39, _
Optional ByVal STARTING_QUARTER As Long = 1, _
Optional ByVal NO_PERIODS As Long = 10, _
Optional ByVal BETA0_VAL As Double = 5, _
Optional ByVal BETA1_VAL As Double = 0.1, _
Optional ByVal OLS_TREND As Boolean = False, _
Optional ByVal WINTER_FACTOR As Double = 0.25, _
Optional ByVal SPRING_FACTOR As Double = -0.15, _
Optional ByVal SUMMER_FACTOR As Double = 0.15, _
Optional ByVal FALL_FACTOR As Double = -0.25, _
Optional ByVal SIGMA_ERROR As Double = 0.1, _
Optional ByVal SE_VERSION As Integer = 0, _
Optional ByVal OUTPUT As Integer = 0)

'Key concept: Seasonal adjustment of a series yt makes sense if the
'series is composed of a series with no seasonal variation, yt*,
'plus a series that contains nothing but seasonal variation, yts.

'Suppose yt is quarterly seasonal.  The series with no seasonal variation,
'yt*, equals b0 + b1 x Time.

'The quarterly seasonal variation values are given in the Optional
'parameters (which you can change).

'The actual, observed quarterly values are yt* + yts + et, which is
'distributed normally with mean zero and SD Error (a controllable
'parameter).

Dim h As Long
Dim i As Long
Dim j As Long
'Winter/Spring/Summer/Fall/Overall
Dim k(1 To 5) As Double
Dim l As Long
Dim SUM_ARR(1 To 5) As Double

Dim FACTOR_VAL As Double 'Mean Factor

Dim TEMP1_MATRIX As Variant
Dim TEMP2_MATRIX As Variant

Dim OLS_MATRIX As Variant
Dim SEASONAL_VECTOR As Variant
Dim DUMMY_VECTOR As Variant 'For regress series on Seasonal Dummies

Dim INDEX0_VECTOR As Variant 'Unadjusted
Dim INDEX1_VECTOR As Variant 'Adjusted

On Error GoTo ERROR_LABEL

'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------
For j = 1 To 5
    k(j) = 0: SUM_ARR(j) = 0
Next j
FACTOR_VAL = (WINTER_FACTOR + SPRING_FACTOR + SUMMER_FACTOR + FALL_FACTOR) / 4
'-------------------------------------------------------------------------------------
ReDim DUMMY_VECTOR(1 To NO_PERIODS + 1, 1 To 4)
ReDim TEMP1_MATRIX(0 To NO_PERIODS + 1, 1 To 5)
'-------------------------------------------------------------------------------------
TEMP1_MATRIX(0, 1) = "INDEX"
TEMP1_MATRIX(0, 2) = "QUARTER"
TEMP1_MATRIX(0, 3) = "NO SEASONAL VARIATION [YT*]"
TEMP1_MATRIX(0, 4) = "SEASONAL VARIATION [YTS]"
TEMP1_MATRIX(0, 5) = "OBSERVED VALUES [YT]"
'-------------------------------------------------------------------------------------
h = STARTING_QUARTER
If h < 1 Then: h = 1
If h > 4 Then: h = 4
l = STARTING_PERIOD
For i = 1 To NO_PERIODS + 1
    TEMP1_MATRIX(i, 1) = l
    DUMMY_VECTOR(i, 1) = TEMP1_MATRIX(i, 1)
    
    TEMP1_MATRIX(i, 2) = h
    TEMP1_MATRIX(i, 3) = BETA0_VAL + BETA1_VAL * TEMP1_MATRIX(i, 1) + RANDOM_NORMAL_FUNC(0, SIGMA_ERROR, 0)

    Select Case h
    Case 1 'Winter
        TEMP1_MATRIX(i, 4) = WINTER_FACTOR
        DUMMY_VECTOR(i, 2) = 1
    Case 2 'Spring
        TEMP1_MATRIX(i, 4) = SPRING_FACTOR
        DUMMY_VECTOR(i, 3) = 1
    Case 3 'Summer
        TEMP1_MATRIX(i, 4) = SUMMER_FACTOR
        DUMMY_VECTOR(i, 4) = 1
    Case 4 'Fall
        TEMP1_MATRIX(i, 4) = FALL_FACTOR
    End Select
    TEMP1_MATRIX(i, 5) = TEMP1_MATRIX(i, 3) + TEMP1_MATRIX(i, 4)
    Select Case h
    Case 1
        SUM_ARR(1) = SUM_ARR(1) + TEMP1_MATRIX(i, 5)
        k(1) = k(1) + 1
    Case 2
        SUM_ARR(2) = SUM_ARR(2) + TEMP1_MATRIX(i, 5)
        k(2) = k(2) + 1
    Case 3
        SUM_ARR(3) = SUM_ARR(3) + TEMP1_MATRIX(i, 5)
        k(3) = k(3) + 1
    Case 4
        SUM_ARR(4) = SUM_ARR(4) + TEMP1_MATRIX(i, 5)
        k(4) = k(4) + 1
    End Select
    SUM_ARR(5) = SUM_ARR(5) + TEMP1_MATRIX(i, 5)
    k(5) = k(5) + 1
    h = h + 1
    If h > 4 Then: h = 1
    l = l + 1
Next i
'-------------------------------------------------------------------------------------
If OLS_TREND = True Then 'Data Generation Process (with Trend)
'-------------------------------------------------------------------------------------
    ReDim INDEX0_VECTOR(1 To NO_PERIODS + 1, 1 To 4)
    ReDim INDEX1_VECTOR(1 To NO_PERIODS + 1, 1 To 1)
    For i = 1 To NO_PERIODS + 1
        For j = 1 To 4: INDEX0_VECTOR(i, j) = DUMMY_VECTOR(i, j): Next j
        INDEX1_VECTOR(i, 1) = TEMP1_MATRIX(i, 5)
    Next i
    OLS_MATRIX = REGRESSION_LS1_FUNC(INDEX0_VECTOR, INDEX1_VECTOR, True, SE_VERSION, 0)

    OLS_MATRIX(6, 1) = "INTERCEPT"
    OLS_MATRIX(7, 1) = "TREND"
    OLS_MATRIX(8, 1) = "WINTER"
    OLS_MATRIX(9, 1) = "SPRING"
    OLS_MATRIX(10, 1) = "SUMMER"
'-------------------------------------------------------------------------------------
Else 'Ordinary Method: Regression without Trend
'-------------------------------------------------------------------------------------
    ReDim INDEX0_VECTOR(1 To NO_PERIODS + 1, 1 To 3)
    ReDim INDEX1_VECTOR(1 To NO_PERIODS + 1, 1 To 1)
    For i = 1 To NO_PERIODS + 1
        For j = 1 To 3: INDEX0_VECTOR(i, j) = DUMMY_VECTOR(i, j + 1): Next j
        INDEX1_VECTOR(i, 1) = TEMP1_MATRIX(i, 5)
    Next i
    OLS_MATRIX = REGRESSION_LS1_FUNC(INDEX0_VECTOR, INDEX1_VECTOR, True, SE_VERSION, 0)
    OLS_MATRIX(6, 1) = "INTERCEPT"
    OLS_MATRIX(7, 1) = "WINTER"
    OLS_MATRIX(8, 1) = "SPRING"
    OLS_MATRIX(9, 1) = "SUMMER"
'-------------------------------------------------------------------------------------
End If
'-------------------------------------------------------------------------------------
ReDim SEASONAL_VECTOR(1 To 5, 1 To 2)
SEASONAL_VECTOR(1, 1) = "WINTER"
SEASONAL_VECTOR(2, 1) = "SPRING"
SEASONAL_VECTOR(3, 1) = "SUMMER"
SEASONAL_VECTOR(4, 1) = "FALL"
SEASONAL_VECTOR(5, 1) = "OVERALL"
For j = 1 To 5: SEASONAL_VECTOR(j, 2) = SUM_ARR(j) / k(j): Next j
'-------------------------------------------------------------------------------------
ReDim INDEX0_VECTOR(1 To 6, 1 To 7) 'Create Seasonal Index
INDEX0_VECTOR(1, 1) = "QUARTER"
INDEX0_VECTOR(2, 1) = 1
INDEX0_VECTOR(3, 1) = 2
INDEX0_VECTOR(4, 1) = 3
INDEX0_VECTOR(5, 1) = 4
INDEX0_VECTOR(6, 1) = ""
'-------------------------------------------------------------------------------------
INDEX0_VECTOR(1, 2) = "SEASON"
INDEX0_VECTOR(2, 2) = "WINTER"
INDEX0_VECTOR(3, 2) = "SPRING"
INDEX0_VECTOR(4, 2) = "SUMMER"
INDEX0_VECTOR(5, 2) = "FALL"
INDEX0_VECTOR(6, 2) = "AVERAGE"
'-------------------------------------------------------------------------------------
INDEX0_VECTOR(1, 3) = "PREDICTED"
'-------------------------------------------------------------------------------------
If OLS_TREND = True Then 'Regression with Trend
'-------------------------------------------------------------------------------------
    INDEX0_VECTOR(2, 3) = OLS_MATRIX(6, 2) + OLS_MATRIX(8, 2)
    INDEX0_VECTOR(3, 3) = OLS_MATRIX(6, 2) + OLS_MATRIX(9, 2)
    INDEX0_VECTOR(4, 3) = OLS_MATRIX(6, 2) + OLS_MATRIX(10, 2)
'-------------------------------------------------------------------------------------
Else
'-------------------------------------------------------------------------------------
    INDEX0_VECTOR(2, 3) = OLS_MATRIX(6, 2) + OLS_MATRIX(7, 2)
    INDEX0_VECTOR(3, 3) = OLS_MATRIX(6, 2) + OLS_MATRIX(8, 2)
    INDEX0_VECTOR(4, 3) = OLS_MATRIX(6, 2) + OLS_MATRIX(9, 2)
'-------------------------------------------------------------------------------------
End If 'Ordinary Method: Regression without Trend
'-------------------------------------------------------------------------------------
INDEX0_VECTOR(5, 3) = OLS_MATRIX(6, 2)
INDEX0_VECTOR(6, 3) = (INDEX0_VECTOR(2, 3) + INDEX0_VECTOR(3, 3) + INDEX0_VECTOR(4, 3) + INDEX0_VECTOR(5, 3)) / 4
'-------------------------------------------------------------------------------------
INDEX0_VECTOR(1, 4) = "SEASONAL INDEX"
INDEX0_VECTOR(2, 4) = INDEX0_VECTOR(2, 3) - INDEX0_VECTOR(6, 3)
INDEX0_VECTOR(3, 4) = INDEX0_VECTOR(3, 3) - INDEX0_VECTOR(6, 3)
INDEX0_VECTOR(4, 4) = INDEX0_VECTOR(4, 3) - INDEX0_VECTOR(6, 3)
INDEX0_VECTOR(5, 4) = INDEX0_VECTOR(5, 3) - INDEX0_VECTOR(6, 3)
INDEX0_VECTOR(6, 4) = ""
'-------------------------------------------------------------------------------------
INDEX0_VECTOR(1, 5) = "ACTUAL INDEX"
INDEX0_VECTOR(2, 5) = WINTER_FACTOR - FACTOR_VAL
INDEX0_VECTOR(3, 5) = SPRING_FACTOR - FACTOR_VAL
INDEX0_VECTOR(4, 5) = SUMMER_FACTOR - FACTOR_VAL
INDEX0_VECTOR(5, 5) = FALL_FACTOR - FACTOR_VAL
INDEX0_VECTOR(6, 5) = ""
'-------------------------------------------------------------------------------------
INDEX0_VECTOR(1, 6) = "PARAMETERS"
INDEX0_VECTOR(2, 6) = WINTER_FACTOR
INDEX0_VECTOR(3, 6) = SPRING_FACTOR
INDEX0_VECTOR(4, 6) = SUMMER_FACTOR
INDEX0_VECTOR(5, 6) = FALL_FACTOR
INDEX0_VECTOR(6, 6) = ""
'-------------------------------------------------------------------------------------
INDEX0_VECTOR(1, 7) = "ACTUAL"
INDEX0_VECTOR(2, 7) = BETA0_VAL + INDEX0_VECTOR(2, 6)
INDEX0_VECTOR(3, 7) = BETA0_VAL + INDEX0_VECTOR(3, 6)
INDEX0_VECTOR(4, 7) = BETA0_VAL + INDEX0_VECTOR(4, 6)
INDEX0_VECTOR(5, 7) = BETA0_VAL + INDEX0_VECTOR(5, 6)
INDEX0_VECTOR(6, 7) = ""
'-------------------------Adjust Series with Seasonal Index---------------------
For j = 1 To 5
    k(j) = 0: SUM_ARR(j) = 0
Next j
'-------------------------------------------------------------------------------------
ReDim TEMP2_MATRIX(0 To NO_PERIODS + 1, 1 To 5) 'Adjust Series with Seasonal Index
TEMP2_MATRIX(0, 1) = "INDEX"
TEMP2_MATRIX(0, 2) = "QUARTER"
TEMP2_MATRIX(0, 3) = "OBSERVED Y"
If OLS_TREND = True Then
    TEMP2_MATRIX(0, 4) = "SEASONAL INDEX ACCOUNTING FOR TREND"
    TEMP2_MATRIX(0, 5) = "SEASONAL AND TREND ADJUSTED Y"
Else
    TEMP2_MATRIX(0, 4) = "SEASONAL INDEX"
    TEMP2_MATRIX(0, 5) = "SEASONALLY ADJUSTED Y"
End If
'-------------------------------------------------------------------------------------
For i = 1 To NO_PERIODS + 1
'-------------------------------------------------------------------------------------
    TEMP2_MATRIX(i, 1) = TEMP1_MATRIX(i, 1)
    TEMP2_MATRIX(i, 2) = TEMP1_MATRIX(i, 2)
    TEMP2_MATRIX(i, 3) = TEMP1_MATRIX(i, 5)
    Select Case TEMP2_MATRIX(i, 2)
    Case 1
        TEMP2_MATRIX(i, 4) = INDEX0_VECTOR(2, 4)
        TEMP2_MATRIX(i, 5) = TEMP2_MATRIX(i, 3) - INDEX0_VECTOR(2, 4)
    Case 2
        TEMP2_MATRIX(i, 4) = INDEX0_VECTOR(3, 4)
        TEMP2_MATRIX(i, 5) = TEMP2_MATRIX(i, 3) - INDEX0_VECTOR(3, 4)
    Case 3
        TEMP2_MATRIX(i, 4) = INDEX0_VECTOR(4, 4)
        TEMP2_MATRIX(i, 5) = TEMP2_MATRIX(i, 3) - INDEX0_VECTOR(4, 4)
    Case 4
        TEMP2_MATRIX(i, 4) = INDEX0_VECTOR(5, 4)
        TEMP2_MATRIX(i, 5) = TEMP2_MATRIX(i, 3) - INDEX0_VECTOR(5, 4)
    End Select
    Select Case TEMP2_MATRIX(i, 2)
    Case 1
        SUM_ARR(1) = SUM_ARR(1) + TEMP2_MATRIX(i, 5)
        k(1) = k(1) + 1
    Case 2
        SUM_ARR(2) = SUM_ARR(2) + TEMP2_MATRIX(i, 5)
        k(2) = k(2) + 1
    Case 3
        SUM_ARR(3) = SUM_ARR(3) + TEMP2_MATRIX(i, 5)
        k(3) = k(3) + 1
    Case 4
        SUM_ARR(4) = SUM_ARR(4) + TEMP2_MATRIX(i, 5)
        k(4) = k(4) + 1
    End Select
    SUM_ARR(5) = SUM_ARR(5) + TEMP2_MATRIX(i, 5)
    k(5) = k(5) + 1
'-------------------------------------------------------------------------------------
Next i
'-------------------------------------------------------------------------------------
ReDim INDEX1_VECTOR(1 To 6, 1 To 3) 'Create Seasonal Index
'-------------------------------------------------------------------------------------
INDEX1_VECTOR(1, 1) = ""
INDEX1_VECTOR(1, 2) = "UNADJUSTED AVERAGE"
INDEX1_VECTOR(1, 3) = "ADJUSTED AVERAGE"
INDEX1_VECTOR(2, 1) = "WINTER"
INDEX1_VECTOR(3, 1) = "SPRING"
INDEX1_VECTOR(4, 1) = "SUMMER"
INDEX1_VECTOR(5, 1) = "FALL"
INDEX1_VECTOR(6, 1) = "OVERALL"
'-------------------------------------------------------------------------------------
For j = 1 To 5
    INDEX1_VECTOR(j + 1, 2) = SEASONAL_VECTOR(j, 2)
    INDEX1_VECTOR(j + 1, 3) = SUM_ARR(j) / k(j)
Next j
'-------------------------------------------------------------------------------------
Select Case OUTPUT
'-------------------------------------------------------------------------------------
Case 0 'Data Generation Process (with Trend)
'-------------------------------------------------------------------------------------
    QUARTERLY_SEASONALITY_SIMULATION_FUNC = TEMP1_MATRIX
'-------------------------------------------------------------------------------------
Case 1 'Series on seasonal dummies
'-------------------------------------------------------------------------------------
    QUARTERLY_SEASONALITY_SIMULATION_FUNC = DUMMY_VECTOR
'-------------------------------------------------------------------------------------
Case 2 'Observed Values Seasonal Vector
'-------------------------------------------------------------------------------------
    QUARTERLY_SEASONALITY_SIMULATION_FUNC = SEASONAL_VECTOR
'-------------------------------------------------------------------------------------
Case 3 'Regress series on seasonal dummies
'-------------------------------------------------------------------------------------
    QUARTERLY_SEASONALITY_SIMULATION_FUNC = OLS_MATRIX
'-------------------------------------------------------------------------------------
Case 4 'Seasonal Index Table
'-------------------------------------------------------------------------------------
    QUARTERLY_SEASONALITY_SIMULATION_FUNC = INDEX0_VECTOR
'-------------------------------------------------------------------------------------
Case 5 'Adjust Series with Seasonal Index
'-------------------------------------------------------------------------------------
    QUARTERLY_SEASONALITY_SIMULATION_FUNC = TEMP2_MATRIX
'-------------------------------------------------------------------------------------
Case 6 'Create Seasonal Index
'-------------------------------------------------------------------------------------
    QUARTERLY_SEASONALITY_SIMULATION_FUNC = INDEX1_VECTOR
'-------------------------------------------------------------------------------------
Case Else
'-------------------------------------------------------------------------------------
    QUARTERLY_SEASONALITY_SIMULATION_FUNC = Array(TEMP1_MATRIX, DUMMY_VECTOR, SEASONAL_VECTOR, OLS_MATRIX, INDEX0_VECTOR, TEMP2_MATRIX, INDEX1_VECTOR)
'-------------------------------------------------------------------------------------
End Select
'-------------------------------------------------------------------------------------
Exit Function
'-------------------------------------------------------------------------------------
ERROR_LABEL:
'-------------------------------------------------------------------------------------
QUARTERLY_SEASONALITY_SIMULATION_FUNC = Err.number
End Function
'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------
Function QUARTERLY_SEASONALITY_FITTING_FUNC(ByRef XDATA_RNG As Variant, _
ByRef YDATA_RNG As Variant, _
ByRef PARAM_RNG As Variant, _
Optional ByRef OPTIMIZER_FLAG As Boolean = False, _
Optional ByVal nLOOPS As Long = 1000, _
Optional ByVal tolerance As Double = 10 ^ -15, _
Optional ByVal epsilon As Double = 10 ^ -10)

Dim YDATA_VECTOR As Variant
Dim XDATA_MATRIX As Variant
Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL

XDATA_MATRIX = XDATA_RNG
If UBound(XDATA_MATRIX, 2) < 2 Then: GoTo ERROR_LABEL 'Period/Quarter/Error
YDATA_VECTOR = YDATA_RNG
If UBound(YDATA_VECTOR, 1) = 1 Then
    YDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(YDATA_VECTOR)
End If
If UBound(YDATA_VECTOR, 1) <> UBound(XDATA_MATRIX, 1) Then: GoTo ERROR_LABEL

PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then
    PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)
End If
If UBound(PARAM_VECTOR, 1) <> 6 Then: GoTo ERROR_LABEL
'[1:6] b0/b1/Winter/Spring/Summer/Fall
If OPTIMIZER_FLAG = True Then
    QUARTERLY_SEASONALITY_FITTING_FUNC = LEVENBERG_MARQUARDT_OPTIMIZATION1_FUNC(XDATA_MATRIX, YDATA_VECTOR, PARAM_VECTOR, "QUARTERLY_SEASONALITY_OBJ_FUNC", "", 0, nLOOPS, tolerance, epsilon)
    'QUARTERLY_SEASONALITY_FITTING_FUNC = LEVENBERG_MARQUARDT_OPTIMIZATION2_FUNC(XDATA_MATRIX, YDATA_VECTOR, PARAM_VECTOR, "QUARTERLY_SEASONALITY_OBJ_FUNC", "", 0, nLOOPS, tolerance, EPSILON)
    PUB_XDATA_MATRIX = XDATA_MATRIX
    PUB_YDATA_VECTOR = YDATA_VECTOR
    'QUARTERLY_SEASONALITY_FITTING_FUNC = NELDER_MEAD_OPTIMIZATION3_FUNC("QUARTERLY_SEASONALITY_OBJ_FUNC", PARAM_VECTOR, nLOOPS, tolerance)
Else
    QUARTERLY_SEASONALITY_FITTING_FUNC = QUARTERLY_SEASONALITY_OBJ_FUNC(XDATA_MATRIX, PARAM_VECTOR)
End If

Exit Function
ERROR_LABEL:
QUARTERLY_SEASONALITY_FITTING_FUNC = Err.number
End Function
'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------
Function QUARTERLY_SEASONALITY_OBJ_FUNC(ByRef XDATA_MATRIX As Variant, _
Optional ByRef PARAM_VECTOR As Variant)

Dim i As Long
Dim j As Long
Dim NROWS As Long

Dim ERROR_FLAG As Boolean
Dim YFIT_VAL As Double
Dim YDATA_VECTOR As Variant
Dim RMS_ERROR As Double

On Error GoTo ERROR_LABEL

ERROR_FLAG = False
If UBound(XDATA_MATRIX, 2) < 2 Then: GoTo ERROR_LABEL
If UBound(XDATA_MATRIX, 2) = 3 Then: ERROR_FLAG = True
NROWS = UBound(XDATA_MATRIX, 1)

If IsArray(PARAM_VECTOR) = True Then
    ReDim YDATA_VECTOR(1 To NROWS, 1 To 1)
    For i = 1 To NROWS
        GoSub CALC_LINE: YDATA_VECTOR(i, 1) = YFIT_VAL
    Next i
    QUARTERLY_SEASONALITY_OBJ_FUNC = YDATA_VECTOR
Else
    PARAM_VECTOR = XDATA_MATRIX
    XDATA_MATRIX = PUB_XDATA_MATRIX
    RMS_ERROR = 0
    For i = 1 To NROWS
        GoSub CALC_LINE
        RMS_ERROR = RMS_ERROR + Abs(PUB_YDATA_VECTOR(i, 1) - YFIT_VAL) ^ 2
    Next i
    QUARTERLY_SEASONALITY_OBJ_FUNC = (RMS_ERROR / NROWS) ^ 0.5
End If

Exit Function
'-----------------------------------------------------------------------------------------------------------
CALC_LINE:
'-----------------------------------------------------------------------------------------------------------
    YFIT_VAL = PARAM_VECTOR(1, 1) + PARAM_VECTOR(2, 1) * XDATA_MATRIX(i, 1) 'Period
    If ERROR_FLAG = True Then: YFIT_VAL = YFIT_VAL + XDATA_MATRIX(i, 3)
    j = XDATA_MATRIX(i, 2) 'Quarter
    YFIT_VAL = YFIT_VAL + PARAM_VECTOR(2 + j, 1)
'-----------------------------------------------------------------------------------------------------------
Return
'-----------------------------------------------------------------------------------------------------------
ERROR_LABEL:
QUARTERLY_SEASONALITY_OBJ_FUNC = PUB_EPSILON
End Function
'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------


