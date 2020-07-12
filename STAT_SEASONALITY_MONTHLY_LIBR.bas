Attribute VB_Name = "STAT_SEASONALITY_MONTHLY_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : MONTHLY_SEASONAL_ANALYSIS_CDM_FUNC
'DESCRIPTION   : Seasonal Analysis - Classical Decomposition Method
'Evaluating Trends with Monthly Cycle data
'Seasonal Analysis of this data using Classical Decomposition Method
'1. Calculate Centered Moving Average
'2. Establish Raw Seasonal Indexes
'3. Establish Normalized Seasonal Indexes
'4. Calculate Deseasonalized values
'5. Establish Deseasonalized Trend Regression
'6. Use Regression Model to Project Deseasonalized Trend
'7. Use Seasonalize Indexes to Project Seasonalized Model Trend
'8. Asses Overall Seasonal Model Fit.

'LIBRARY       : STATISTICS
'GROUP         : SEASONALITY
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'REFERENCE     :
'************************************************************************************
'************************************************************************************


Function MONTHLY_SEASONAL_ANALYSIS_CDM_FUNC(ByRef DATE_RNG As Variant, _
ByRef DATA_RNG As Variant, _
Optional ByRef PARAM_RNG As Variant, _
Optional ByVal OUTPUT As Integer = 0)

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long
Dim n As Long
Dim NROWS As Long

Dim YEAR_INT As Long
Dim MONTH_INT As Long
Dim DATE_VAL As Date
Dim RMS_ERROR As Double

Dim DATA_MATRIX As Variant 'Original Data/Moving Avgs/Decomposition Factors/Model
Dim INDEXES_MATRIX As Variant 'Seasonal Indexes by Month
Dim YDATA_VECTOR As Variant
Dim XDATA_MATRIX As Variant
Dim COEF_VECTOR As Variant

Dim DATA_VECTOR As Variant
Dim DATE_VECTOR As Variant

Dim INDEX_OBJ As New Collection

On Error GoTo ERROR_LABEL

'On Error Resume Next

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then
    DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
End If

DATE_VECTOR = DATE_RNG
If UBound(DATE_VECTOR, 1) = 1 Then
    DATE_VECTOR = MATRIX_TRANSPOSE_FUNC(DATE_VECTOR)
End If
If UBound(DATA_VECTOR, 1) <> UBound(DATE_VECTOR, 1) Then: GoTo ERROR_LABEL
NROWS = UBound(DATA_VECTOR, 1)
If NROWS < 12 Then: GoTo ERROR_LABEL
ReDim DATA_MATRIX(0 To NROWS, 1 To 12)

DATA_MATRIX(0, 1) = "DATE"
DATA_MATRIX(0, 2) = "DATA"
DATA_MATRIX(0, 3) = "MA12"
DATA_MATRIX(0, 4) = "CENTERED MA12"
DATA_MATRIX(0, 5) = "RAW SEAS DIFF"
DATA_MATRIX(0, 6) = "RAW SEAS RATIO"
DATA_MATRIX(0, 7) = "NOR SEAS INDEX"
DATA_MATRIX(0, 8) = "DESEASONALIZED DATA"
DATA_MATRIX(0, 9) = "DESEASONALIZED TREND"
DATA_MATRIX(0, 10) = "SEASONAL MODEL PREDICT"
DATA_MATRIX(0, 11) = "SEASONAL DIFF"
DATA_MATRIX(0, 12) = "RESIDUAL"

'----------------------------------------------------------------------------
i = 1: l = 1
YEAR_INT = Year(DATE_VECTOR(i, 1))
Call INDEX_OBJ.Add(i, CStr(l))
For i = 1 To 6
    GoSub INDEX_LINE
    RMS_ERROR = RMS_ERROR + DATA_VECTOR(i, 1)
    RMS_ERROR = RMS_ERROR + DATA_VECTOR(i + 6, 1)
    'For j = 3 To 6: DATA_MATRIX(i, j) = "": Next j
Next i
DATA_MATRIX(i - 1, 3) = RMS_ERROR / 12
'----------------------------------------------------------------------------
For i = 7 To NROWS
    GoSub INDEX_LINE
    RMS_ERROR = RMS_ERROR - DATA_VECTOR(i - 6, 1)
    If i <= NROWS - 6 Then
        RMS_ERROR = RMS_ERROR + DATA_VECTOR(i + 6, 1)
        DATA_MATRIX(i, 3) = RMS_ERROR / 12
        j = 1
    Else
        DATA_MATRIX(i, 3) = RMS_ERROR / (12 - j)
        j = j + 1
    End If
    DATA_MATRIX(i, 4) = 0.5 * (DATA_MATRIX(i - 1, 3) + DATA_MATRIX(i, 3))
    DATA_MATRIX(i, 5) = DATA_MATRIX(i, 2) - DATA_MATRIX(i, 4)
    If DATA_MATRIX(i, 4) <> 0 Then
        DATA_MATRIX(i, 6) = DATA_MATRIX(i, 2) / DATA_MATRIX(i, 4)
    End If
Next i
GoSub MONTHLY_LINE
If OUTPUT = 1 Then
    MONTHLY_SEASONAL_ANALYSIS_CDM_FUNC = INDEXES_MATRIX
    Exit Function
End If

If IsArray(PARAM_RNG) = True Then
    COEF_VECTOR = PARAM_RNG
    If UBound(COEF_VECTOR, 1) = 1 Then
        COEF_VECTOR = MATRIX_TRANSPOSE_FUNC(COEF_VECTOR)
    End If
    If UBound(COEF_VECTOR, 1) <> 3 Then: GoTo ERROR_LABEL
    For i = 1 To NROWS
        MONTH_INT = Month(DATA_MATRIX(i, 1))
        DATA_MATRIX(i, 7) = INDEXES_MATRIX(l + 6, 1 + MONTH_INT)
        If DATA_MATRIX(i, 7) <> 0 Then
            DATA_MATRIX(i, 8) = DATA_MATRIX(i, 2) / DATA_MATRIX(i, 7)
        End If
        DATA_MATRIX(i, 9) = (COEF_VECTOR(1, 1) + COEF_VECTOR(2, 1) * i + COEF_VECTOR(3, 1) * i ^ 2)
        DATA_MATRIX(i, 10) = DATA_MATRIX(i, 9) * DATA_MATRIX(i, 7)
        DATA_MATRIX(i, 11) = DATA_MATRIX(i, 10) - DATA_MATRIX(i, 9)
        DATA_MATRIX(i, 12) = DATA_MATRIX(i, 2) - DATA_MATRIX(i, 10)
    Next i
Else
    ReDim YDATA_VECTOR(1 To NROWS, 1 To 1)
    ReDim XDATA_MATRIX(1 To NROWS, 1 To 1) '2)
    
    For i = 1 To NROWS
        MONTH_INT = Month(DATA_MATRIX(i, 1))
        DATA_MATRIX(i, 7) = INDEXES_MATRIX(l + 6, 1 + MONTH_INT)
        If DATA_MATRIX(i, 7) <> 0 Then
            DATA_MATRIX(i, 8) = DATA_MATRIX(i, 2) / DATA_MATRIX(i, 7)
        End If
        YDATA_VECTOR(i, 1) = DATA_MATRIX(i, 8)
        XDATA_MATRIX(i, 1) = i ': XDATA_MATRIX(i, 2) = i ^ 2
    Next i
    'Deseasonalized Trend Regression Analysis: Y = a + b * Month_No + c * Month_No ^ 2
    'COEF_VECTOR = REGRESSION_MULT_COEF_FUNC(XDATA_MATRIX, YDATA_VECTOR, True, 2)
    COEF_VECTOR = POLYNOMIAL_REGRESSION_FUNC(XDATA_MATRIX, YDATA_VECTOR, 2, 0)
    If OUTPUT = 2 Then
        MONTHLY_SEASONAL_ANALYSIS_CDM_FUNC = COEF_VECTOR
        Exit Function
    End If
    For i = 1 To NROWS
        DATA_MATRIX(i, 9) = (COEF_VECTOR(1, 1) + COEF_VECTOR(2, 1) * i + COEF_VECTOR(3, 1) * i ^ 2)
        DATA_MATRIX(i, 10) = DATA_MATRIX(i, 9) * DATA_MATRIX(i, 7)
        DATA_MATRIX(i, 11) = DATA_MATRIX(i, 10) - DATA_MATRIX(i, 9)
        DATA_MATRIX(i, 12) = DATA_MATRIX(i, 2) - DATA_MATRIX(i, 10)
    Next i
End If

If OUTPUT = 0 Then
    MONTHLY_SEASONAL_ANALYSIS_CDM_FUNC = DATA_MATRIX
Else
    MONTHLY_SEASONAL_ANALYSIS_CDM_FUNC = Array(DATA_MATRIX, INDEXES_MATRIX)
End If

Exit Function
'----------------------------------------------------------------------------
INDEX_LINE:
'----------------------------------------------------------------------------
    DATA_MATRIX(i, 1) = DATE_VECTOR(i, 1)
    DATA_MATRIX(i, 2) = DATA_VECTOR(i, 1)
    n = Year(DATE_VECTOR(i, 1))
    If n <> YEAR_INT Then
        l = l + 1
        Call INDEX_OBJ.Add(i, CStr(l))
        YEAR_INT = n
    End If
'----------------------------------------------------------------------------
Return
'----------------------------------------------------------------------------
MONTHLY_LINE:
'Procedure to transform monthlly data table to database
'Source data is organized with months across and years down
'----------------------------------------------------------------------------
ReDim INDEXES_MATRIX(1 To l + 7, 1 To 12 + 2)
DATE_VAL = Now
INDEXES_MATRIX(1, 1) = "YEAR"
For h = 1 To 12
    INDEXES_MATRIX(1, 1 + h) = Format(DateSerial(Year(DATE_VAL), h, 1), "mmm")
    INDEXES_MATRIX(l + 4, 1 + h) = -2 ^ 52
    INDEXES_MATRIX(l + 5, 1 + h) = 2 ^ 52
Next h
INDEXES_MATRIX(1, h + 1) = "IDX"
INDEXES_MATRIX(l + 2, 1) = "SUM"
INDEXES_MATRIX(l + 3, 1) = "COUNTA"
INDEXES_MATRIX(l + 4, 1) = "MAX"
INDEXES_MATRIX(l + 5, 1) = "MIN"
INDEXES_MATRIX(l + 6, 1) = "TRIM"
INDEXES_MATRIX(l + 7, 1) = "NOR"

For h = 1 To l
    i = CLng(INDEX_OBJ.Item(h))
    INDEXES_MATRIX(1 + h, 1) = Year(DATE_VECTOR(i, 1))
    If h <> l Then
        j = CLng(INDEX_OBJ.Item(h + 1)) - 1
    Else
        j = NROWS
    End If
    For k = i To j 'For each month per year
        MONTH_INT = Month(DATE_VECTOR(k, 1))
        If DATA_MATRIX(k, 6) <> "" And DATA_MATRIX(k, 6) <> 0 Then
            INDEXES_MATRIX(1 + h, 1 + MONTH_INT) = DATA_MATRIX(k, 6)
            INDEXES_MATRIX(1 + h, 2 + 12) = INDEXES_MATRIX(1 + h, 2 + 12) + DATA_MATRIX(k, 6)
            
            INDEXES_MATRIX(l + 2, 1 + MONTH_INT) = INDEXES_MATRIX(l + 2, 1 + MONTH_INT) + DATA_MATRIX(k, 6)
            INDEXES_MATRIX(l + 2, 2 + 12) = INDEXES_MATRIX(l + 2, 2 + 12) + DATA_MATRIX(k, 6)
            
            INDEXES_MATRIX(l + 3, 1 + MONTH_INT) = INDEXES_MATRIX(l + 3, 1 + MONTH_INT) + 1
            INDEXES_MATRIX(l + 3, 2 + 12) = INDEXES_MATRIX(l + 3, 2 + 12) + 1
            
            If DATA_MATRIX(k, 6) > INDEXES_MATRIX(l + 4, 1 + MONTH_INT) Then: INDEXES_MATRIX(l + 4, 1 + MONTH_INT) = DATA_MATRIX(k, 6)
            If DATA_MATRIX(k, 6) < INDEXES_MATRIX(l + 5, 1 + MONTH_INT) Then: INDEXES_MATRIX(l + 5, 1 + MONTH_INT) = DATA_MATRIX(k, 6)
        End If
    Next k
Next h

INDEXES_MATRIX(l + 4, 2 + 12) = "" '-2 ^ 52
INDEXES_MATRIX(l + 5, 2 + 12) = "" '2 ^ 52
For h = 1 To 12
    'If INDEXES_MATRIX(l + 4, 1 + h) > INDEXES_MATRIX(l + 4, 2 + 12) Then: INDEXES_MATRIX(l + 4, 2 + 12) = INDEXES_MATRIX(l + 4, 1 + h)
    'If INDEXES_MATRIX(l + 5, 1 + h) < INDEXES_MATRIX(l + 5, 2 + 12) Then: INDEXES_MATRIX(l + 5, 2 + 12) = INDEXES_MATRIX(l + 5, 1 + h)
    If (INDEXES_MATRIX(l + 3, 1 + h) - 2) <> 0 Then
        INDEXES_MATRIX(l + 6, 1 + h) = (INDEXES_MATRIX(l + 2, 1 + h) - INDEXES_MATRIX(l + 4, 1 + h) - INDEXES_MATRIX(l + 5, 1 + h)) / (INDEXES_MATRIX(l + 3, 1 + h) - 2)
        INDEXES_MATRIX(l + 6, 2 + 12) = INDEXES_MATRIX(l + 6, 2 + 12) + INDEXES_MATRIX(l + 6, 1 + h)
    End If
Next h
For h = 1 To 12
    If INDEXES_MATRIX(l + 6, 2 + 12) <> 0 Then
        INDEXES_MATRIX(l + 7, 1 + h) = INDEXES_MATRIX(l + 6, 1 + h) * 12 / INDEXES_MATRIX(l + 6, 2 + 12)
        INDEXES_MATRIX(l + 7, 2 + 12) = INDEXES_MATRIX(l + 7, 2 + 12) + INDEXES_MATRIX(l + 7, 1 + h)
    End If
Next h
'----------------------------------------------------------------------------
Return
ERROR_LABEL:
MONTHLY_SEASONAL_ANALYSIS_CDM_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MONTHLY_SEASONALITY_INDEX_FUNC
'DESCRIPTION   : 'This functions performs the first step in creating the
'seasonal index: regress the unadjusted series on monthly dummies.
'Then it creates the Seasonal Index using the predicted values from the
'Regression.
'LIBRARY       : STATISTICS
'GROUP         : SEASONALITY
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'REFERENCE     : http://www.bls.gov/cps/cpsatabs.htm
'************************************************************************************
'************************************************************************************

Function MONTHLY_SEASONALITY_INDEX_FUNC(ByRef DATE_RNG As Variant, _
ByRef UNADJUSTED_RNG As Variant, _
Optional ByRef ADJUSTED_RNG As Variant, _
Optional ByVal INTERCEPT_FLAG As Boolean = True, _
Optional ByVal SE_VERSION As Integer = 0, _
Optional ByVal OUTPUT As Integer = 0)

Dim i As Long
Dim j As Long
Dim NROWS As Long

Dim RMS_ERROR As Double
Dim FACTOR_VAL As Double
Dim TEMP_VECTOR As Variant
Dim TEMP_MATRIX As Variant
Dim SEASONAL_VECTOR As Variant 'Seasonal Index

Dim DATE_VECTOR As Variant
Dim ADJUSTED_VECTOR As Variant 'Seasonally Adjusted DataSet
Dim UNADJUSTED_VECTOR As Variant 'Not Seasonally Adjusted DataSet
Dim OLS_MATRIX As Variant 'We use these estimates to create the Seasonal Index

On Error GoTo ERROR_LABEL

DATE_VECTOR = DATE_RNG
If UBound(DATE_VECTOR, 1) = 1 Then
    DATE_VECTOR = MATRIX_TRANSPOSE_FUNC(DATE_VECTOR)
End If

UNADJUSTED_VECTOR = UNADJUSTED_RNG
If UBound(UNADJUSTED_VECTOR, 1) = 1 Then
    UNADJUSTED_VECTOR = MATRIX_TRANSPOSE_FUNC(UNADJUSTED_VECTOR)
End If
NROWS = UBound(DATE_VECTOR, 1)
ReDim SEASONAL_VECTOR(0 To 12, 1 To 3)
    
SEASONAL_VECTOR(0, 1) = "MONTH"
SEASONAL_VECTOR(0, 2) = "PREDICTED"
SEASONAL_VECTOR(0, 3) = "INDEX"

For i = 1 To 12: SEASONAL_VECTOR(i, 1) = Format(DateSerial(Year(Date), i, 1), "mmm"): Next i

RMS_ERROR = 0
'---------------------------------------------------------------------------------------------
If INTERCEPT_FLAG = True Then
'---------------------------------------------------------------------------------------------
    ReDim TEMP_MATRIX(1 To NROWS, 1 To 11) 'Unadjusted Data on Monthly Dummies
    For i = 1 To NROWS
        j = Month(DATE_VECTOR(i, 1))
        If j >= 1 And j <= 11 Then: TEMP_MATRIX(i, j) = 1
    Next i
    'Regression of Unadjusted Data on Monthly Dummies
    OLS_MATRIX = REGRESSION_LS1_FUNC(TEMP_MATRIX, UNADJUSTED_VECTOR, INTERCEPT_FLAG, SE_VERSION, 0)
    For i = 1 To 11: OLS_MATRIX(i + 6, 1) = SEASONAL_VECTOR(i, 1): Next i
    RMS_ERROR = 0
    For i = 1 To 12
        If i <> 12 Then
            SEASONAL_VECTOR(i, 2) = OLS_MATRIX(6 + i, 2) + OLS_MATRIX(6, 2)
        Else
            SEASONAL_VECTOR(i, 2) = OLS_MATRIX(6, 2)
        End If
        RMS_ERROR = RMS_ERROR + SEASONAL_VECTOR(i, 2)
    Next i
    For i = 1 To 12: SEASONAL_VECTOR(i, 3) = SEASONAL_VECTOR(i, 2) - RMS_ERROR / 12: Next i
'---------------------------------------------------------------------------------------------
ElseIf INTERCEPT_FLAG = False Then
'---------------------------------------------------------------------------------------------
    ReDim TEMP_MATRIX(1 To NROWS, 1 To 12) 'Unadjusted Data on Monthly Dummies
    For i = 1 To NROWS
        For j = 1 To 12
            If Month(DATE_VECTOR(i, 1)) = j Then: TEMP_MATRIX(i, j) = 1
        Next j
    Next i
    'Regression of Unadjusted Data on Monthly Dummies
    OLS_MATRIX = REGRESSION_LS1_FUNC(TEMP_MATRIX, UNADJUSTED_VECTOR, INTERCEPT_FLAG, SE_VERSION, 0)
    For i = 1 To 12
        OLS_MATRIX(i + 5, 1) = SEASONAL_VECTOR(i, 1)
        SEASONAL_VECTOR(i, 2) = OLS_MATRIX(5 + i, 2)
        RMS_ERROR = RMS_ERROR + SEASONAL_VECTOR(i, 2)
    Next i
    For i = 1 To 12: SEASONAL_VECTOR(i, 3) = SEASONAL_VECTOR(i, 2) - RMS_ERROR / 12: Next i
'---------------------------------------------------------------------------------------------
End If
'---------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------
Select Case OUTPUT
'--------------------------------------------------------------------------------------------
Case 0
'performs the first step in creating the seasonal index: regress the unadjusted series on monthly dummies
'--------------------------------------------------------------------------------------------
    MONTHLY_SEASONALITY_INDEX_FUNC = OLS_MATRIX
'--------------------------------------------------------------------------------------------
Case 1
'creates the Seasonal Index using the predicted values from the Regression
'--------------------------------------------------------------------------------------------
    MONTHLY_SEASONALITY_INDEX_FUNC = SEASONAL_VECTOR
'--------------------------------------------------------------------------------------------
Case Else 'Regression Adjusted
'compares the Seasonally Adjusted series with the official Seasonally Adjusted Index.
'The regression-based method should bot quite far from the method used by the stat office,
'and the discrepancies shall be smaller in more recent years.
'--------------------------------------------------------------------------------------------
    ADJUSTED_VECTOR = ADJUSTED_RNG
    If UBound(ADJUSTED_VECTOR, 1) = 1 Then
        ADJUSTED_VECTOR = MATRIX_TRANSPOSE_FUNC(ADJUSTED_VECTOR)
    End If
    If UBound(ADJUSTED_VECTOR, 1) <> NROWS Then: GoTo ERROR_LABEL

    ReDim TEMP_VECTOR(0 To NROWS, 1 To 7)
    TEMP_VECTOR(0, 1) = "DATE"
    TEMP_VECTOR(0, 2) = "UNADJUSTED"
    TEMP_VECTOR(0, 3) = "BLS ADJUSTED"
    TEMP_VECTOR(0, 4) = "REGRESSION ADJUSTED"
    TEMP_VECTOR(0, 5) = "BLS ADJUSTED - REGRESSION ADJUSTED"
    TEMP_VECTOR(0, 6) = "BLS ADJUSTED - UNADJUSTED"
    TEMP_VECTOR(0, 7) = "REGRESSION ADJUSTED - UNADJUSTED"
    For i = 1 To NROWS
        TEMP_VECTOR(i, 1) = DATE_VECTOR(i, 1)
        TEMP_VECTOR(i, 2) = UNADJUSTED_VECTOR(i, 1)
        TEMP_VECTOR(i, 3) = ADJUSTED_VECTOR(i, 1)
'demonstrates the effect of seasonal adjustment on the unemployment series.
        j = Month(DATE_VECTOR(i, 1))
        If j >= 1 And j <= 12 Then: FACTOR_VAL = SEASONAL_VECTOR(j, 3)
        TEMP_VECTOR(i, 4) = UNADJUSTED_VECTOR(i, 1) - FACTOR_VAL
        'TEMP_VECTOR(i, 4) = WorksheetFunction.Round(UNADJUSTED_VECTOR(i, 1) - FACTOR_VAL, 1)
        TEMP_VECTOR(i, 5) = TEMP_VECTOR(i, 3) - TEMP_VECTOR(i, 4)
        TEMP_VECTOR(i, 6) = TEMP_VECTOR(i, 3) - TEMP_VECTOR(i, 2)
        TEMP_VECTOR(i, 7) = TEMP_VECTOR(i, 4) - TEMP_VECTOR(i, 2)
    Next i
    If OUTPUT = 2 Then
        MONTHLY_SEASONALITY_INDEX_FUNC = TEMP_VECTOR
    Else
        MONTHLY_SEASONALITY_INDEX_FUNC = Array(OLS_MATRIX, SEASONAL_VECTOR, TEMP_VECTOR)
    End If
'--------------------------------------------------------------------------------------------
End Select
'--------------------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
MONTHLY_SEASONALITY_INDEX_FUNC = Err.number
End Function
