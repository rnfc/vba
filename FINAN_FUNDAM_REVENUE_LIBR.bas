Attribute VB_Name = "FINAN_FUNDAM_REVENUE_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : REVENUE_CAGR_FUNC

'DESCRIPTION   : RETURNS EXPECTED REVENUE AND COMPOUNDED ANNUALIZED GROWTH
'RATES DEPENDING ON THE MARKET SHARE

'LIBRARY       : FUNDAMENTAL
'GROUP         : REVENUE
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 03/02/2008
'************************************************************************************
'************************************************************************************

Function REVENUE_CAGR_FUNC(ByVal CURRENT_REVENUE As Double, _
ByVal MARKET_SIZE As Double, _
ByVal GROWTH_RATE_MARKET As Double, _
ByRef DATA_SHARES_RNG As Variant, _
Optional ByVal OUTPUT As Integer = 0)

Dim i As Long
Dim j As Long

Dim SROW As Long
Dim NROWS As Long

Dim SCOLUMN As Long
Dim NCOLUMNS As Long

Dim TENOR_VAL As Double

Dim TEMP1_MATRIX As Variant
Dim TEMP2_MATRIX As Variant

Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_SHARES_RNG
'COLUMN 1 FOR DATES_SHARES_RNG - MUST BE DATES OR TENORS (YEAR_END)
'COLUMN 2 FOR DATES_SHARES_RNG - MUST BE YOUR EXPECTED MARKET SHARE

NROWS = UBound(DATA_MATRIX, 1)
    
ReDim TEMP1_MATRIX(1 To NROWS + 1, 1 To 3)
TEMP1_MATRIX(1, 1) = "YEAR"
TEMP1_MATRIX(1, 2) = "EXPECTED REVENUES"
TEMP1_MATRIX(1, 3) = "E[CAGR]"

For i = 2 To NROWS + 1
'FROM Settlement Date to Date i
     TEMP1_MATRIX(i, 1) = DATA_MATRIX(i - 1, 1)
     TEMP1_MATRIX(i, 2) = MARKET_SIZE * (1 + GROWTH_RATE_MARKET) ^ DATA_MATRIX(i - 1, 1) * DATA_MATRIX(i - 1, 2) 'Expected revenues in year Date(i)
     TEMP1_MATRIX(i, 3) = (TEMP1_MATRIX(i, 2) / CURRENT_REVENUE) ^ (1 / DATA_MATRIX(i - 1, 1)) - 1
     'Expected compounded growth rate; From Settlement to Date(i)
Next i

ReDim TEMP2_MATRIX(1 To NROWS, 1 To NROWS)
'RELATIVE ANALYSIS: FROM Date j to Date i
TEMP2_MATRIX(1, 1) = "E[CAGR]"
For j = 2 To NROWS
   TEMP2_MATRIX(1, j) = DATA_MATRIX(j - 1, 1)
   TEMP2_MATRIX(j, 1) = DATA_MATRIX(j, 1)
   TENOR_VAL = DATA_MATRIX(j - 1, 1) + 1
   For i = j To NROWS
        TEMP2_MATRIX(i, j) = (TEMP1_MATRIX(i + 1, 2) / TEMP1_MATRIX(j, 2)) ^ (1 / (1 + DATA_MATRIX(i, 1) - TENOR_VAL)) - 1
   Next i
Next j

SROW = LBound(TEMP2_MATRIX, 1): NROWS = UBound(TEMP2_MATRIX, 1)
SCOLUMN = LBound(TEMP2_MATRIX, 2): NCOLUMNS = UBound(TEMP2_MATRIX, 2)

For j = SCOLUMN To NCOLUMNS
    For i = SROW To NROWS
        If TEMP2_MATRIX(i, j) = 0 Then: TEMP2_MATRIX(i, j) = ""
    Next i
Next j

Select Case OUTPUT
Case 0
    REVENUE_CAGR_FUNC = TEMP2_MATRIX ' RELATIVE ANALYSIS
Case 1
    REVENUE_CAGR_FUNC = TEMP1_MATRIX ' CAGR MATRIX
Case Else
    REVENUE_CAGR_FUNC = Array(TEMP1_MATRIX, TEMP2_MATRIX)
End Select

Exit Function
ERROR_LABEL:
REVENUE_CAGR_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : REVENUE_FORECASTING_FUNC
'DESCRIPTION   : The following function takes a set of data on two variables,
'such as prices and corresponding demands, and then estimates the
'best-fitting linear, exponential, and power curves for these data.
'It also calculates the corresponding mean absolute percentage error (MAPE).
'LIBRARY       : FUNDAMENTAL
'GROUP         : REVENUE
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 03/02/2008
'************************************************************************************
'************************************************************************************

Function REVENUE_FORECASTING_FUNC(ByRef YDATA_RNG As Variant, _
ByRef XDATA_RNG As Variant, _
Optional ByVal OUTPUT As Integer = 0)

Dim i As Long
Dim NROWS As Long

Dim LINEAR_SUM As Double
Dim POWER_SUM As Double
Dim EXPON_SUM As Double

Dim YLOG_VECTOR As Variant
Dim XLOG_VECTOR As Variant

Dim YDATA_VECTOR As Variant
Dim XDATA_VECTOR As Variant

Dim RESULT_MATRIX As Variant
Dim MAPE_MATRIX As Variant

On Error GoTo ERROR_LABEL

YDATA_VECTOR = YDATA_RNG
If UBound(YDATA_VECTOR, 1) = 1 Then
    YDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(YDATA_VECTOR)
End If

XDATA_VECTOR = XDATA_RNG
If UBound(XDATA_VECTOR, 1) = 1 Then
    XDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(XDATA_VECTOR)
End If
If UBound(YDATA_VECTOR, 1) <> UBound(XDATA_VECTOR, 1) Then: GoTo ERROR_LABEL
NROWS = UBound(YDATA_VECTOR, 1)

ReDim YLOG_VECTOR(1 To NROWS, 1 To 1)
ReDim XLOG_VECTOR(1 To NROWS, 1 To 1)

For i = 1 To NROWS
    YLOG_VECTOR(i, 1) = Log(YDATA_VECTOR(i, 1))
    XLOG_VECTOR(i, 1) = Log(XDATA_VECTOR(i, 1))
Next i

ReDim RESULT_MATRIX(1 To 3, 1 To 3)
'LINEAR
RESULT_MATRIX(1, 1) = REGRESSION_SIMPLE_COEF_FUNC(XDATA_VECTOR, YDATA_VECTOR)(2, 1) 'INTERCEPT
RESULT_MATRIX(2, 1) = REGRESSION_SIMPLE_COEF_FUNC(XDATA_VECTOR, YDATA_VECTOR)(1, 1) 'SLOPE

'A logarithmic trendline is a best-fit curved line that is most
'useful when the rate of change in the data increases or decreases
'quickly and then levels out. A logarithmic trendline can use
'negative and/or positive values.

'POWER
RESULT_MATRIX(1, 2) = Exp(REGRESSION_SIMPLE_COEF_FUNC(XLOG_VECTOR, YLOG_VECTOR)(2, 1))
RESULT_MATRIX(2, 2) = REGRESSION_SIMPLE_COEF_FUNC(XLOG_VECTOR, YLOG_VECTOR)(1, 1)

'A power trendline is a curved line that is best used with data sets
'that compare measurements that increase at a specific rate — for
'example, the acceleration of a race car at 1-second intervals. You
'cannot create a power trendline if your data contains zero or negative values.

'EXPONENTIAL
RESULT_MATRIX(1, 3) = Exp(REGRESSION_SIMPLE_COEF_FUNC(XDATA_VECTOR, YLOG_VECTOR)(2, 1))
RESULT_MATRIX(2, 3) = REGRESSION_SIMPLE_COEF_FUNC(XDATA_VECTOR, YLOG_VECTOR)(1, 1)

'An exponential trendline is a curved line that is most useful
'when data values rise or fall at increasingly higher rates. You
'cannot create an exponential trendline if your data contains
'zero or negative values.

ReDim MAPE_MATRIX(1 To NROWS, 1 To 3)

For i = 1 To NROWS
    MAPE_MATRIX(i, 1) = Abs(YDATA_VECTOR(i, 1) - (RESULT_MATRIX(1, 1) + RESULT_MATRIX(2, 1) * XDATA_VECTOR(i, 1))) / YDATA_VECTOR(i, 1) 'LINEAR
    LINEAR_SUM = LINEAR_SUM + MAPE_MATRIX(i, 1)
    
    MAPE_MATRIX(i, 2) = Abs(YDATA_VECTOR(i, 1) - (RESULT_MATRIX(1, 2) * XDATA_VECTOR(i, 1) ^ RESULT_MATRIX(2, 2))) / YDATA_VECTOR(i, 1) 'POWER
    POWER_SUM = POWER_SUM + MAPE_MATRIX(i, 2)
    
    MAPE_MATRIX(i, 3) = Abs(YDATA_VECTOR(i, 1) - (RESULT_MATRIX(1, 3) * _
    Exp(RESULT_MATRIX(2, 3) * XDATA_VECTOR(i, 1)))) / YDATA_VECTOR(i, 1) 'EXPONENTIAL
    
    EXPON_SUM = EXPON_SUM + MAPE_MATRIX(i, 3)
Next i

RESULT_MATRIX(3, 1) = LINEAR_SUM / NROWS 'MAPE
RESULT_MATRIX(3, 2) = POWER_SUM / NROWS 'MAPE
RESULT_MATRIX(3, 3) = EXPON_SUM / NROWS 'MAPE

Select Case OUTPUT
Case 0
    REVENUE_FORECASTING_FUNC = RESULT_MATRIX
Case Else
    REVENUE_FORECASTING_FUNC = MAPE_MATRIX
End Select

Exit Function
ERROR_LABEL:
REVENUE_FORECASTING_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : REVENUE_FORECASTING_CONTROL_FUNC
'DESCRIPTION   : Use specific measurements to track and control revenue forecasting
'LIBRARY       : FUNDAMENTAL
'GROUP         : REVENUE
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 03/02/2008
'************************************************************************************
'************************************************************************************

Function REVENUE_FORECASTING_CONTROL_FUNC(ByRef TENOR_RNG As Variant, _
ByRef REVENUE_RNG As Variant, _
ByRef ESTIMATED_RNG As Variant, _
Optional ByVal OUTPUT As Integer = 1)

Dim i As Long
Dim NROWS As Long

Dim TENOR_VECTOR As Variant
Dim REVENUE_VECTOR As Variant
Dim ESTIMATED_VECTOR As Variant

Dim TEMP_SUM As Double
Dim TEMP_MATRIX As Variant

Dim MEAN_ABS_ERROR As Double 'Mean Absolute Error - An absolute value of
'forecast errors, does not place weight on the amount of the error. Calculated
'as the sum of (actual values - predicted values) / n.

Dim MEAN_SQRT_ERROR As Double 'Mean Square Error - Similar to Mean Absolute
'Error, but does place more emphasis on the amount of error; i.e. an error of
'8 is twice as significant as 4. Calculated as the sum of (actual values -
'predicted values)^2 / n.

Dim ROOT_MEAN_SQRT_ERROR As Double 'Root Mean Square Error - To make the
'Mean Square Error useful and comparable to the Mean Absolute Error, we
'can take the square root of the Mean Square Error. We can then use this
'as a guide to establish an error limit or standard for flagging
'unacceptable errors.

On Error GoTo ERROR_LABEL

REVENUE_VECTOR = REVENUE_RNG
If UBound(REVENUE_VECTOR, 1) = 1 Then
    REVENUE_VECTOR = MATRIX_TRANSPOSE_FUNC(REVENUE_VECTOR)
End If

ESTIMATED_VECTOR = ESTIMATED_RNG
If UBound(ESTIMATED_VECTOR, 1) = 1 Then
    ESTIMATED_VECTOR = MATRIX_TRANSPOSE_FUNC(ESTIMATED_VECTOR)
End If

TENOR_VECTOR = TENOR_RNG
If UBound(TENOR_VECTOR, 1) = 1 Then
    TENOR_VECTOR = MATRIX_TRANSPOSE_FUNC(TENOR_VECTOR)
End If

NROWS = UBound(REVENUE_VECTOR, 1)

If UBound(REVENUE_VECTOR, 1) <> UBound(ESTIMATED_VECTOR, 1) Then: GoTo ERROR_LABEL
If UBound(REVENUE_VECTOR, 1) <> UBound(TENOR_VECTOR, 1) Then: GoTo ERROR_LABEL

ReDim TEMP_MATRIX(0 To NROWS + 1, 1 To 6)

TEMP_MATRIX(0, 1) = "TN"
TEMP_MATRIX(0, 2) = "ACTUAL"
TEMP_MATRIX(0, 3) = "ESTIMATION"
TEMP_MATRIX(0, 4) = "ERROR"

TEMP_SUM = 0
For i = 1 To NROWS
    TEMP_MATRIX(i, 1) = TENOR_VECTOR(i, 1)
    TEMP_MATRIX(i, 2) = REVENUE_VECTOR(i, 1)
    TEMP_MATRIX(i, 3) = ESTIMATED_VECTOR(i, 1)
    TEMP_MATRIX(i, 4) = TEMP_MATRIX(i, 2) - TEMP_MATRIX(i, 3)
    TEMP_SUM = TEMP_SUM + TEMP_MATRIX(i, 4)
    TEMP_MATRIX(i, 5) = Abs(TEMP_MATRIX(i, 4))
    TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 5) ^ 2
Next i

TEMP_MATRIX(NROWS + 1, 1) = ""
TEMP_MATRIX(NROWS + 1, 2) = ""
TEMP_MATRIX(NROWS + 1, 3) = ("SUM")
TEMP_MATRIX(NROWS + 1, 4) = TEMP_SUM
TEMP_MATRIX(NROWS + 1, 5) = Abs(TEMP_MATRIX(NROWS + 1, 4))
TEMP_MATRIX(NROWS + 1, 6) = TEMP_MATRIX(NROWS + 1, 5) ^ 2


MEAN_ABS_ERROR = TEMP_MATRIX(NROWS + 1, 5) / NROWS
TEMP_MATRIX(0, 5) = "ABSOLUTE: " & Format(MEAN_ABS_ERROR, "0.00")

MEAN_SQRT_ERROR = TEMP_MATRIX(NROWS + 1, 6) / NROWS
ROOT_MEAN_SQRT_ERROR = MEAN_SQRT_ERROR ^ 0.5

TEMP_MATRIX(0, 6) = "ERROR_SQRT: " & Format(ROOT_MEAN_SQRT_ERROR, "0.00")

Select Case OUTPUT
Case 0
    REVENUE_FORECASTING_CONTROL_FUNC = TEMP_MATRIX
Case Else
    ReDim TEMP_MATRIX(1 To 3, 1 To 2)

    TEMP_MATRIX(1, 1) = "MEAN ABSOLUTE ERROR"
    TEMP_MATRIX(1, 2) = MEAN_ABS_ERROR
    
    'Setting a reasonable limit requires a review of past history and other
    'factors. However, we can begin by looking at the Root Mean Square Error
    'as a starting point.
    
    TEMP_MATRIX(2, 1) = "MEAN SQUARE ERROR"
    TEMP_MATRIX(2, 2) = MEAN_SQRT_ERROR
    
    TEMP_MATRIX(3, 1) = "ROOT MEAN SQUARE ERROR"
    TEMP_MATRIX(3, 2) = ROOT_MEAN_SQRT_ERROR
    
    REVENUE_FORECASTING_CONTROL_FUNC = TEMP_MATRIX
End Select

Exit Function
ERROR_LABEL:
REVENUE_FORECASTING_CONTROL_FUNC = Err.number
End Function
