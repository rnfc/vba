Attribute VB_Name = "FINAN_ASSET_TA_SORNETTE_LIBR"

'----------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'----------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------
Private Const PUB_TDAYS_PER_YEAR As Long = 252
Private PUB_MIN_MAX_INT As Integer
Private PUB_LAST_LOG_VAL As Double
Private PUB_XDATA_VECTOR As Variant
Private PUB_YDATA_VECTOR As Variant
Private Const PUB_EPSILON As Double = 2 ^ 52 '1E-100

'----------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------
'Market Crashed & Wave Theory
'----------------------------------------------------------------------------------------------
'Large financial crashes
'Didier Sornette and Anders Johansen
'http://arxiv.org/PS_cache/cond-mat/pdf/9704/9704127v2.pdf
'http://www.ess.ucla.edu/faculty/sornette/
'http://www.gummy-stuff.org/sornette.htm
'----------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------

Function ASSET_WAVES_SORNETTE_FUNC(ByVal TICKER_STR As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal END_DATE As Date, _
Optional ByRef PARAM_RNG As Variant, _
Optional ByVal NO_PERIODS As Long = 10, _
Optional ByRef HOLIDAYS_RNG As Variant, _
Optional ByVal OUTPUT As Integer = 0)

Dim i As Long
Dim j As Long
Dim NROWS As Long
Dim NCOLUMNS As Long
Dim TEMP_ARR As Variant
Dim RMS_ERROR As Double

Dim PARAM_VECTOR As Variant
Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

If IsArray(TICKER_STR) = False Then
    DATA_MATRIX = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, START_DATE, END_DATE, _
                  "d", "DA", False, True, True)
Else
    DATA_MATRIX = TICKER_STR
End If
NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then
    PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)
End If

ReDim TEMP_MATRIX(0 To NROWS + NO_PERIODS, 1 To 9)
TEMP_MATRIX(0, 1) = "DATE"
TEMP_MATRIX(0, 2) = "PRICE"
TEMP_MATRIX(0, 3) = "TT"
TEMP_MATRIX(0, 4) = "E(TT)"
TEMP_MATRIX(0, 5) = "OSC(TT)"
TEMP_MATRIX(0, 6) = "SORNETTE(TT)"
TEMP_MATRIX(0, 7) = "LOG(PRICE)"
TEMP_MATRIX(0, 9) = "FORECASTING"

RMS_ERROR = 0
For i = 1 To NROWS - 1
    TEMP_MATRIX(i, 1) = DATA_MATRIX(i, 1)
    TEMP_MATRIX(i, 2) = DATA_MATRIX(i, 2)
    START_DATE = TEMP_MATRIX(i, 1)
    END_DATE = DATA_MATRIX(NROWS, 1)
    GoSub SORNETTE_LINE
    TEMP_MATRIX(i, 7) = Log(TEMP_MATRIX(i, 2))
    TEMP_MATRIX(i, 8) = Abs(TEMP_MATRIX(i, 6) - TEMP_MATRIX(i, 7))
    RMS_ERROR = RMS_ERROR + TEMP_MATRIX(i, 8) ^ 2
Next i
RMS_ERROR = (RMS_ERROR / (NROWS - 1)) ^ 0.5
If OUTPUT <> 0 Then
    ASSET_WAVES_SORNETTE_FUNC = RMS_ERROR
    Exit Function
End If

TEMP_MATRIX(0, 8) = "RMS ERROR: " & Format(RMS_ERROR, "0.00%")
i = NROWS
TEMP_MATRIX(i, 1) = DATA_MATRIX(i, 1)
TEMP_MATRIX(i, 2) = DATA_MATRIX(i, 2)
For j = 7 To 9: TEMP_MATRIX(i, j) = "": Next j

START_DATE = TEMP_MATRIX(NROWS, 1)
For i = NROWS + 1 To NROWS + NO_PERIODS
    TEMP_MATRIX(i, 1) = WORKDAY2_FUNC(TEMP_MATRIX(i - 1, 1), 1, HOLIDAYS_RNG)
    TEMP_MATRIX(i, 2) = ""
    END_DATE = TEMP_MATRIX(i, 1)
    GoSub SORNETTE_LINE
    For j = 7 To 8: TEMP_MATRIX(i, j) = "": Next j
    TEMP_MATRIX(i, 9) = Exp(TEMP_MATRIX(i, 6))
Next i
    
ASSET_WAVES_SORNETTE_FUNC = TEMP_MATRIX

'---------------------------------------------------------------------------------
Exit Function
'---------------------------------------------------------------------------------
SORNETTE_LINE:
'---------------------------------------------------------------------------------
TEMP_ARR = ASSET_WAVES_SORNETTE_PRICING_FUNC(PARAM_VECTOR, START_DATE, END_DATE, 1)
TEMP_MATRIX(i, 3) = TEMP_ARR(LBound(TEMP_ARR) + 0)
TEMP_MATRIX(i, 4) = TEMP_ARR(LBound(TEMP_ARR) + 1)
TEMP_MATRIX(i, 5) = TEMP_ARR(LBound(TEMP_ARR) + 2)
TEMP_MATRIX(i, 6) = TEMP_ARR(LBound(TEMP_ARR) + 3)
Erase TEMP_ARR
'---------------------------------------------------------------------------------
Return
'---------------------------------------------------------------------------------
ERROR_LABEL:
ASSET_WAVES_SORNETTE_FUNC = PUB_EPSILON
End Function

Function ASSET_WAVES_SORNETTE_OPTIMIZER_FUNC(ByRef DATES_RNG As Variant, _
ByRef PRICES_RNG As Variant, _
ByRef PARAM_RNG As Variant, _
Optional ByVal VERSION As Integer = 0)

Dim i As Long
Dim NROWS As Long
Dim PARAM_VECTOR As Variant
Dim XDATA_VECTOR  As Variant
Dim YDATA_VECTOR  As Variant

On Error GoTo ERROR_LABEL

PARAM_VECTOR = PARAM_RNG 'prices
If UBound(PARAM_VECTOR, 1) = 1 Then
    PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)
End If

YDATA_VECTOR = PRICES_RNG 'prices
If UBound(YDATA_VECTOR, 1) = 1 Then
    YDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(YDATA_VECTOR)
End If

XDATA_VECTOR = DATES_RNG 'dates
If UBound(XDATA_VECTOR, 1) = 1 Then
    XDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(XDATA_VECTOR)
End If
If UBound(XDATA_VECTOR, 1) <> UBound(YDATA_VECTOR, 1) Then: GoTo ERROR_LABEL
NROWS = UBound(XDATA_VECTOR, 1)

PUB_XDATA_VECTOR = XDATA_VECTOR
PUB_YDATA_VECTOR = YDATA_VECTOR
'-------------------------------------------------------------------------------------------
Select Case VERSION
'-------------------------------------------------------------------------------------------
Case 0
'-------------------------------------------------------------------------------------------
    PUB_MIN_MAX_INT = 1
    PARAM_VECTOR = NELDER_MEAD_OPTIMIZATION4_FUNC("ASSET_WAVES_SORNETTE_ERROR_FUNC", XDATA_VECTOR, _
                   YDATA_VECTOR, PARAM_VECTOR, PARAM_VECTOR)

'    PARAM_VECTOR = _
    NELDER_MEAD_OPTIMIZATION2_FUNC("ASSET_WAVES_SORNETTE_ERROR_FUNC", XDATA_VECTOR, PARAM_VECTOR)
'-------------------------------------------------------------------------------------------
Case 1
'-------------------------------------------------------------------------------------------
    PUB_MIN_MAX_INT = 1
    For i = 1 To NROWS: PUB_YDATA_VECTOR(i, 1) = Log(PUB_YDATA_VECTOR(i, 1)): Next i
    PUB_LAST_LOG_VAL = PUB_YDATA_VECTOR(NROWS, 1)
    
    PARAM_VECTOR = LEVENBERG_MARQUARDT_OPTIMIZATION3_FUNC(PUB_XDATA_VECTOR, PUB_YDATA_VECTOR, _
    PARAM_VECTOR, "ASSET_WAVES_SORNETTE_FITNESS_FUNC", "ASSET_WAVES_SORNETTE_JACOBI_FUNC", 1000, 1, 1)
    
'    PARAM_VECTOR = LEVENBERG_MARQUARDT_OPTIMIZATION2_FUNC(PUB_XDATA_VECTOR, PUB_YDATA_VECTOR, _
    PARAM_VECTOR, "ASSET_WAVES_SORNETTE_FITNESS_FUNC", "ASSET_WAVES_SORNETTE_JACOBI_FUNC", 0)
'-------------------------------------------------------------------------------------------
Case Else
'-------------------------------------------------------------------------------------------
    PUB_MIN_MAX_INT = -1
    PARAM_VECTOR = PIKAIA_OPTIMIZATION_FUNC("ASSET_WAVES_SORNETTE_ERROR_FUNC", _
                   PARAM_VECTOR, False, , , , , , , , , , , , , , 0)
'-------------------------------------------------------------------------------------------
End Select
'-------------------------------------------------------------------------------------------

ASSET_WAVES_SORNETTE_OPTIMIZER_FUNC = PARAM_VECTOR

Exit Function
ERROR_LABEL:
ASSET_WAVES_SORNETTE_OPTIMIZER_FUNC = Err.number
End Function

Function ASSET_WAVES_SORNETTE_FITNESS_FUNC(ByRef XDATA_VECTOR As Variant, _
Optional ByRef PARAM_VECTOR As Variant)

Dim i As Long
Dim NROWS As Long
Dim YDATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

If IsArray(PARAM_VECTOR) = False Then: PARAM_VECTOR = XDATA_VECTOR
NROWS = UBound(PUB_XDATA_VECTOR, 1)
ReDim YDATA_VECTOR(1 To NROWS, 1 To 1)

For i = 1 To NROWS - 1
    YDATA_VECTOR(i, 1) = ASSET_WAVES_SORNETTE_PRICING_FUNC(PARAM_VECTOR, _
                         PUB_XDATA_VECTOR(i, 1), PUB_XDATA_VECTOR(NROWS, 1), 0)
Next i
YDATA_VECTOR(NROWS, 1) = PUB_LAST_LOG_VAL
ASSET_WAVES_SORNETTE_FITNESS_FUNC = YDATA_VECTOR

Exit Function
ERROR_LABEL:
ASSET_WAVES_SORNETTE_FITNESS_FUNC = Err.number
End Function

Function ASSET_WAVES_SORNETTE_ERROR_FUNC(ByRef XDATA_VECTOR As Variant, _
Optional ByRef YDATA_RNG As Variant, _
Optional ByRef CONST_RNG As Variant, _
Optional ByRef PARAM_VECTOR As Variant)

Dim i As Long
Dim NROWS As Long
Dim TEMP_VAL As Double
Dim RMS_ERROR As Double

On Error GoTo ERROR_LABEL

If IsArray(PARAM_VECTOR) = False Then: PARAM_VECTOR = XDATA_VECTOR

NROWS = UBound(PUB_XDATA_VECTOR, 1)

For i = 1 To NROWS - 1
    TEMP_VAL = ASSET_WAVES_SORNETTE_PRICING_FUNC(PARAM_VECTOR, PUB_XDATA_VECTOR(i, 1), _
               PUB_XDATA_VECTOR(NROWS, 1), 0)
    TEMP_VAL = Abs(TEMP_VAL - Log(PUB_YDATA_VECTOR(i, 1)))
    'Error Value
    RMS_ERROR = RMS_ERROR + TEMP_VAL ^ 2 'RMS-ERROR
Next i
RMS_ERROR = (RMS_ERROR / (NROWS - 1)) ^ 0.5
ASSET_WAVES_SORNETTE_ERROR_FUNC = RMS_ERROR * PUB_MIN_MAX_INT

Exit Function
ERROR_LABEL:
ASSET_WAVES_SORNETTE_ERROR_FUNC = PUB_EPSILON
End Function

Function ASSET_WAVES_SORNETTE_PRICING_FUNC(ByVal PARAM_RNG As Variant, _
ByVal START_DATE As Date, _
ByVal END_DATE As Date, _
Optional ByVal OUTPUT As Integer = 0)

Dim DELTA_VAL As Double
Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double
Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL

PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then
    PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)
End If

DELTA_VAL = (END_DATE - START_DATE) / PUB_TDAYS_PER_YEAR

ATEMP_VAL = PARAM_VECTOR(2, 1) * (DELTA_VAL ^ _
            PARAM_VECTOR(7, 1)) / Sqr(1 + (DELTA_VAL / _
            PARAM_VECTOR(6, 1)) ^ (2 * PARAM_VECTOR(7, 1)))
                    
BTEMP_VAL = 1 + PARAM_VECTOR(3, 1) * Cos(PARAM_VECTOR(4, 1) * _
            Log(DELTA_VAL) + (PARAM_VECTOR(5, 1) / _
            (2 * PARAM_VECTOR(7, 1))) * Log(1 + (DELTA_VAL _
            / PARAM_VECTOR(6, 1)) ^ (2 * PARAM_VECTOR(7, 1))))

Select Case OUTPUT
Case 0
    ASSET_WAVES_SORNETTE_PRICING_FUNC = PARAM_VECTOR(1, 1) + ATEMP_VAL * BTEMP_VAL
Case Else
    ASSET_WAVES_SORNETTE_PRICING_FUNC = Array(DELTA_VAL, ATEMP_VAL, BTEMP_VAL, _
                                        PARAM_VECTOR(1, 1) + ATEMP_VAL * BTEMP_VAL)
End Select

Exit Function
ERROR_LABEL:
ASSET_WAVES_SORNETTE_PRICING_FUNC = Err.number
End Function

Function ASSET_WAVES_SORNETTE_JACOBI_FUNC(ByRef XDATA_RNG As Variant, _
ByRef PARAM_RNG As Variant)

Dim XDATA_VECTOR As Variant
Dim PARAM_VECTOR As Variant

Dim tolerance As Double

On Error GoTo ERROR_LABEL

tolerance = 0.001
XDATA_VECTOR = XDATA_RNG
If UBound(XDATA_VECTOR, 1) = 1 Then
    XDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(XDATA_VECTOR)
End If

PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then
    PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)
End If

ASSET_WAVES_SORNETTE_JACOBI_FUNC = JACOBI_MATRIX_FUNC("ASSET_WAVES_SORNETTE_FITNESS_FUNC", _
                                   XDATA_VECTOR, PARAM_VECTOR, tolerance)
Exit Function
ERROR_LABEL:
ASSET_WAVES_SORNETTE_JACOBI_FUNC = Err.number
End Function

Function ASSET_WAVES_SORNETTE_GUESS_PARAMETERS_FUNC(ByVal PRICES_RNG As Variant)

Dim i As Long
Dim NROWS As Long

Dim TEMP_VAL As Double
Dim MAX_LOG_PRICE As Double

Dim DATA_VECTOR As Variant
Dim PARAM_VECTOR(1 To 7, 1 To 1)

On Error GoTo ERROR_LABEL

DATA_VECTOR = PRICES_RNG
If UBound(DATA_VECTOR, 1) = 1 Then
    DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
End If
NROWS = UBound(DATA_VECTOR, 1)

'Sornette(t) = A + E(t) Osc(t) = A - B (tc - t)^a / SQRT[1 +
'{ (tc - t)/dt }^2a] [1 + C COS[ w Log(tc-t) +
'(dw/2a) Log(1+{ (tc - t)/dt }^2a) ]]

'There are a bunch of parameters, namely: A, B, C, w, dw, dt and a.
'We pick the parameters so as to mimic the behaviour of the S&P ...
'before the crash. You will have to read Sornette's paper to learn
'what parameters to choose.

'The Sornette curve oscillates just like the S&P did, before the
'crash.

MAX_LOG_PRICE = -2 ^ 52
For i = 1 To NROWS
    TEMP_VAL = Log(DATA_VECTOR(i, 1))
    If TEMP_VAL > MAX_LOG_PRICE Then: MAX_LOG_PRICE = TEMP_VAL
Next i

PARAM_VECTOR(1, 1) = MAX_LOG_PRICE '7.4 / 5.9 --> a
PARAM_VECTOR(2, 1) = -(Int(5 * Rnd) + 35) / 100 '-0.3 / -0.38 --> b
PARAM_VECTOR(3, 1) = (Int(3 * Rnd) + 40) / 1000 '0.1 / 0.113 --> cc
PARAM_VECTOR(4, 1) = (Int(10 * Rnd) + 80) / 10 '8 / 8.7 --> w
PARAM_VECTOR(5, 1) = Int(5 * Rnd) + 6 '20 / 18 --> dw
PARAM_VECTOR(6, 1) = Int(4 * Rnd) + 8 '10 / 11 --> dt
PARAM_VECTOR(7, 1) = (Int(8 * Rnd) + 60) / 100 '0.5 / 0.68 --> alpha

ASSET_WAVES_SORNETTE_GUESS_PARAMETERS_FUNC = PARAM_VECTOR

Exit Function
ERROR_LABEL:
ASSET_WAVES_SORNETTE_GUESS_PARAMETERS_FUNC = Err.number
End Function


Function ASSET_WAVES_SORNETTE_SENSITIVITY_FUNC(ByRef DATA_RNG As Variant, _
ByRef PARAM_RNG As Variant, _
Optional ByVal MIN_A As Double = 7.4, _
Optional ByVal MAX_A As Double = 7.9, _
Optional ByVal DELTA_A As Double = 0.2, _
Optional ByVal MIN_B As Double = -0.4, _
Optional ByVal MAX_B As Double = -0.2, _
Optional ByVal DELTA_B As Double = 0.1, _
Optional ByVal MIN_C As Double = 0.1, _
Optional ByVal MAX_C As Double = 0.4, _
Optional ByVal DELTA_C As Double = 0.1, _
Optional ByVal MIN_W As Double = 6, _
Optional ByVal MAX_W As Double = 10, _
Optional ByVal DELTA_W As Double = 1, _
Optional ByVal MIN_DW As Double = 20, _
Optional ByVal MAX_DW As Double = 50, _
Optional ByVal DELTA_DW As Double = 5, _
Optional ByVal MIN_DT As Double = 6, _
Optional ByVal MAX_DT As Double = 14, _
Optional ByVal DELTA_DT As Double = 2, _
Optional ByVal MIN_ALPHA As Double = 0.3, _
Optional ByVal MAX_ALPHA As Double = 0.5, _
Optional ByVal DELTA_ALPHA As Double = 0.1)

Dim A As Double
Dim B As Double
Dim c As Double
Dim W As Double
Dim DW As Double
Dim dt As Double

Dim ALPHA As Double

Dim TEMP_A As Double
Dim TEMP_B As Double
Dim TEMP_C As Double
Dim TEMP_W As Double
Dim TEMP_DW As Double
Dim TEMP_DT As Double
Dim TEMP_ALPHA As Double

Dim BEST_VAL As Variant
Dim TEMP_VAL As Variant

Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL

TEMP_VAL = ASSET_WAVES_SORNETTE_FUNC(DATA_RNG, , , PARAM_RNG, 2)
BEST_VAL = TEMP_VAL(UBound(TEMP_VAL))

TEMP_A = TEMP_VAL(LBound(TEMP_VAL))(1, 1)
TEMP_B = TEMP_VAL(LBound(TEMP_VAL))(2, 1)
TEMP_C = TEMP_VAL(LBound(TEMP_VAL))(3, 1)
TEMP_W = TEMP_VAL(LBound(TEMP_VAL))(4, 1)
TEMP_DW = TEMP_VAL(LBound(TEMP_VAL))(5, 1)
TEMP_DT = TEMP_VAL(LBound(TEMP_VAL))(6, 1)
TEMP_ALPHA = TEMP_VAL(LBound(TEMP_VAL))(7, 1)

ReDim PARAM_VECTOR(1 To 7, 1 To 1)
For A = MIN_A To MAX_A Step DELTA_A
    For W = MIN_W To MAX_W Step DELTA_W
        For B = MIN_B To MAX_B Step DELTA_B
            For c = MIN_C To MAX_C Step DELTA_C
                For DW = MIN_DW To MAX_DW Step DELTA_DW
                    For dt = MIN_DT To MAX_DT Step DELTA_DT
                        For ALPHA = MIN_ALPHA To MAX_ALPHA Step DELTA_ALPHA
                        
                            PARAM_VECTOR(1, 1) = A '/ 10
                            PARAM_VECTOR(2, 1) = B '/ 100
                            PARAM_VECTOR(3, 1) = c '/ 1000
                            PARAM_VECTOR(4, 1) = W
                            PARAM_VECTOR(5, 1) = DW
                            PARAM_VECTOR(6, 1) = dt
                            PARAM_VECTOR(7, 1) = ALPHA '/ 100

                            TEMP_VAL = ASSET_WAVES_SORNETTE_FUNC(DATA_RNG, , , PARAM_VECTOR, 1)
                            If TEMP_VAL < BEST_VAL Then
                                BEST_VAL = TEMP_VAL
                                TEMP_A = PARAM_VECTOR(1, 1)
                                TEMP_B = PARAM_VECTOR(2, 1)
                                TEMP_C = PARAM_VECTOR(3, 1)
                                TEMP_W = PARAM_VECTOR(4, 1)
                                TEMP_DW = PARAM_VECTOR(5, 1)
                                TEMP_DT = PARAM_VECTOR(6, 1)
                                TEMP_ALPHA = PARAM_VECTOR(7, 1)
                            End If
1983:
                        Next ALPHA
                    Next dt
                Next DW
            Next c
        Next B
    Next W
Next A

PARAM_VECTOR(1, 1) = TEMP_A
PARAM_VECTOR(2, 1) = TEMP_B
PARAM_VECTOR(3, 1) = TEMP_C
PARAM_VECTOR(4, 1) = TEMP_W
PARAM_VECTOR(5, 1) = TEMP_DW
PARAM_VECTOR(6, 1) = TEMP_DT
PARAM_VECTOR(7, 1) = TEMP_ALPHA

ASSET_WAVES_SORNETTE_SENSITIVITY_FUNC = PARAM_VECTOR

Exit Function
ERROR_LABEL:
ASSET_WAVES_SORNETTE_SENSITIVITY_FUNC = Err.number
End Function
