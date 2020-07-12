Attribute VB_Name = "FINAN_ASSET_MOMENTS_ELLIP_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------
'generating ellipses that contain points in a scatter plot
'http://www.gummy-stuff.org/stock-ellipses.htm
'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------

Function ASSET_ELLIPSES_FUNC(ByVal INDEX_STR As String, _
ByVal STOCK_STR As String, _
ByVal START_DATE As Date, _
ByVal END_DATE As Date, _
Optional ByVal ANGLE_VAL As Double = 62, _
Optional ByVal SCALE_VAL As Double = 1.61, _
Optional ByVal OUTPUT As Integer = 0)

'ANGLE_VAL --> Rotate Ellipse
'SCALE_VAL --> reSize the Ellipse

Dim i As Long
Dim NROWS As Long

Dim A_VAL As Double
Dim B_VAL As Double
Dim C_VAL As Double
Dim T_VAL As Double

Dim PI_VAL As Double
Dim COS_VAL As Double
Dim SIN_VAL As Double

Dim TEMP1_SUM As Double
Dim TEMP2_SUM As Double

Dim MEAN1_VAL As Double
Dim MEAN2_VAL As Double

Dim SIGMA1_VAL As Double
Dim SIGMA2_VAL As Double

Dim FACTOR_VAL As Double
Dim SLOPE_VAL As Double
Dim INTERCEPT_VAL As Double
Dim CORRELATION_VAL As Double

Dim XTEMP_VECTOR As Variant
Dim YTEMP_VECTOR As Variant

Dim TEMP1_MATRIX As Variant
Dim TEMP2_MATRIX As Variant

Dim TICKERS_VECTOR As Variant

On Error GoTo ERROR_LABEL

PI_VAL = 3.14159265358979
C_VAL = PI_VAL / 180
FACTOR_VAL = C_VAL * ANGLE_VAL
FACTOR_VAL = ATAN_FUNC(FACTOR_VAL)
COS_VAL = Cos(FACTOR_VAL)
SIN_VAL = Sin(FACTOR_VAL)

ReDim TICKERS_VECTOR(1 To 2, 1 To 1)
TICKERS_VECTOR(1, 1) = INDEX_STR
TICKERS_VECTOR(2, 1) = STOCK_STR

TEMP1_MATRIX = YAHOO_HISTORICAL_DATA_SERIES1_FUNC(TICKERS_VECTOR, START_DATE, END_DATE, 6, "d", False, True)
Erase TICKERS_VECTOR
NROWS = UBound(TEMP1_MATRIX, 1)
ReDim TEMP2_MATRIX(0 To NROWS, 1 To 10)

i = 0
TEMP2_MATRIX(i, 1) = "DATE"
TEMP2_MATRIX(i, 2) = UCase(INDEX_STR) & ": PRICE"
TEMP2_MATRIX(i, 3) = UCase(STOCK_STR) & ": PRICE"

TEMP2_MATRIX(i, 4) = UCase(INDEX_STR) & ": RETURN"
TEMP2_MATRIX(i, 5) = UCase(STOCK_STR) & ": RETURN"
TEMP2_MATRIX(i, 6) = "T"
TEMP2_MATRIX(i, 7) = "U-ELLIPSE"
TEMP2_MATRIX(i, 8) = "V-ELLIPSE"
TEMP2_MATRIX(i, 9) = "U' (" & Format(ANGLE_VAL, "0") & ")"
TEMP2_MATRIX(i, 10) = "V'"

i = 1
TEMP2_MATRIX(i, 1) = TEMP1_MATRIX(i, 1)
TEMP2_MATRIX(i, 2) = TEMP1_MATRIX(i, 2)
TEMP2_MATRIX(i, 3) = TEMP1_MATRIX(i, 3)
TEMP2_MATRIX(i, 4) = ""
TEMP2_MATRIX(i, 5) = ""

TEMP1_SUM = 0: TEMP2_SUM = 0
For i = 2 To NROWS
    TEMP2_MATRIX(i, 1) = TEMP1_MATRIX(i, 1)
    TEMP2_MATRIX(i, 2) = TEMP1_MATRIX(i, 2)
    TEMP2_MATRIX(i, 3) = TEMP1_MATRIX(i, 3)
    
    TEMP2_MATRIX(i, 4) = TEMP1_MATRIX(i, 2) / TEMP1_MATRIX(i - 1, 2) - 1 'Index Return
    TEMP1_SUM = TEMP1_SUM + TEMP2_MATRIX(i, 4)
    TEMP2_MATRIX(i, 5) = TEMP1_MATRIX(i, 3) / TEMP1_MATRIX(i - 1, 3) - 1 'Stock Return
    TEMP2_SUM = TEMP2_SUM + TEMP2_MATRIX(i, 5)
Next i

ReDim XTEMP_VECTOR(1 To NROWS - 1, 1 To 1)
ReDim YTEMP_VECTOR(1 To NROWS - 1, 1 To 1)

MEAN1_VAL = TEMP1_SUM / (NROWS - 1)
MEAN2_VAL = TEMP2_SUM / (NROWS - 1)

TEMP1_SUM = 0: TEMP2_SUM = 0
For i = 2 To NROWS
    TEMP1_SUM = TEMP1_SUM + (TEMP2_MATRIX(i, 4) - MEAN1_VAL) ^ 2
    TEMP2_SUM = TEMP2_SUM + (TEMP2_MATRIX(i, 5) - MEAN2_VAL) ^ 2

    TEMP2_MATRIX(i, 4) = TEMP2_MATRIX(i, 4) - MEAN1_VAL
    XTEMP_VECTOR(i - 1, 1) = TEMP2_MATRIX(i, 4)
    
    TEMP2_MATRIX(i, 5) = TEMP2_MATRIX(i, 5) - MEAN2_VAL
    YTEMP_VECTOR(i - 1, 1) = TEMP2_MATRIX(i, 5)
Next i
SIGMA1_VAL = (TEMP1_SUM / (NROWS - 1)) ^ 0.5
SIGMA2_VAL = (TEMP2_SUM / (NROWS - 1)) ^ 0.5

A_VAL = SIGMA1_VAL * SCALE_VAL
B_VAL = SIGMA2_VAL * SCALE_VAL

TEMP1_MATRIX = REGRESSION_SIMPLE_COEF_FUNC(XTEMP_VECTOR, YTEMP_VECTOR)
SLOPE_VAL = TEMP1_MATRIX(1, 1)
INTERCEPT_VAL = TEMP1_MATRIX(2, 1)

If OUTPUT <> 0 Then
    CORRELATION_VAL = CORRELATION_FUNC(XTEMP_VECTOR, YTEMP_VECTOR, 0, 0)
    ASSET_ELLIPSES_FUNC = Array(SLOPE_VAL, INTERCEPT_VAL, CORRELATION_VAL, _
                          MEAN1_VAL, MEAN2_VAL, SIGMA1_VAL, SIGMA2_VAL)
    Exit Function
End If
Erase XTEMP_VECTOR: Erase YTEMP_VECTOR: Erase TEMP1_MATRIX

i = 1
T_VAL = 0
TEMP2_MATRIX(i, 6) = T_VAL
TEMP2_MATRIX(i, 7) = A_VAL * Cos(C_VAL * T_VAL)
TEMP2_MATRIX(i, 8) = SLOPE_VAL * TEMP2_MATRIX(i, 7) + INTERCEPT_VAL + B_VAL * Sin(C_VAL * T_VAL)
TEMP2_MATRIX(i, 9) = TEMP2_MATRIX(i, 7) * COS_VAL - TEMP2_MATRIX(i, 8) * SIN_VAL
TEMP2_MATRIX(i, 10) = TEMP2_MATRIX(i, 7) * SIN_VAL + TEMP2_MATRIX(i, 8) * COS_VAL

T_VAL = T_VAL + 10
For i = 2 To NROWS
    If T_VAL <= 370 Then
        TEMP2_MATRIX(i, 6) = T_VAL
        TEMP2_MATRIX(i, 7) = A_VAL * Cos(C_VAL * T_VAL)
        TEMP2_MATRIX(i, 8) = SLOPE_VAL * TEMP2_MATRIX(i, 7) + INTERCEPT_VAL + B_VAL * Sin(C_VAL * T_VAL)
        TEMP2_MATRIX(i, 9) = TEMP2_MATRIX(i, 7) * COS_VAL - TEMP2_MATRIX(i, 8) * SIN_VAL
        TEMP2_MATRIX(i, 10) = TEMP2_MATRIX(i, 7) * SIN_VAL + TEMP2_MATRIX(i, 8) * COS_VAL
        T_VAL = T_VAL + 10
    Else
        TEMP2_MATRIX(i, 6) = ""
        TEMP2_MATRIX(i, 7) = ""
        TEMP2_MATRIX(i, 8) = ""
        TEMP2_MATRIX(i, 9) = ""
        TEMP2_MATRIX(i, 10) = ""
    End If
Next i

ASSET_ELLIPSES_FUNC = TEMP2_MATRIX

Exit Function
ERROR_LABEL:
ASSET_ELLIPSES_FUNC = Err.number
End Function
