Attribute VB_Name = "FINAN_ASSET_PROCESS_FIT_LIBR"

'////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////
Private Const PUB_FORMULA_TYPE As Integer = 1
Private Const PUB_EPSILON As Double = 2 ^ 52 '1E-100
Private PUB_WEIGHT_VAL As Double
Private PUB_XDATA_VECTOR As Variant
Private PUB_YDATA_VECTOR As Variant
Private PUB_ERROR_TYPE As Integer
'////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////

Function ASSET_OPEN_PRICE_FITNESS_FUNC(ByVal TICKER_STR As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal END_DATE As Date, _
Optional ByRef PARAM_RNG As Variant, _
Optional ByVal OPTIM_TYPE As Integer = 0, _
Optional ByVal ERROR_TYPE As Integer = 2, _
Optional ByVal CONFIDENCE_VAL As Double = 0.01, _
Optional ByVal WEIGHT_VAL As Double = 0.9, _
Optional ByVal OUTPUT As Integer = 1)

Dim i As Long

Dim NROWS As Long

Dim RMS_ERROR As Double 'Error
Dim MAX_ERROR As Double 'Max Error
Dim AVG_ERROR As Double 'AVG Error
Dim WGT_ERROR As Double 'd Error
Dim TEMP_SUM As Double 'weight factor

Dim DATA_MATRIX As Variant
Dim TEMP_MATRIX As Variant
Dim XDATA_VECTOR As Variant
Dim YDATA_VECTOR As Variant
Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL

If IsArray(TICKER_STR) = False Then
    DATA_MATRIX = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, START_DATE, END_DATE, _
                  "d", "DOC", False, True, True)
Else
    DATA_MATRIX = TICKER_STR
End If
NROWS = UBound(DATA_MATRIX, 1)
PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then
    PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)
End If

'---------------------------------------------------------------------------------------------------
If OUTPUT = 0 Then
'---------------------------------------------------------------------------------------------------
    ReDim YDATA_VECTOR(1 To NROWS - 1, 1 To 1)
    ReDim XDATA_VECTOR(1 To NROWS - 1, 1 To 1)
    
    For i = 1 To NROWS - 1
        YDATA_VECTOR(i, 1) = DATA_MATRIX(i + 1, 2) 'Open(t+1)
        XDATA_VECTOR(i, 1) = DATA_MATRIX(i, 3) 'close(t)
    Next i
    
    ASSET_OPEN_PRICE_FITNESS_FUNC = ASSET_OPEN_PRICE_OPTIMIZER_FUNC(XDATA_VECTOR, YDATA_VECTOR, _
                                    PARAM_VECTOR, WEIGHT_VAL, OPTIM_TYPE, ERROR_TYPE, 0)
    Exit Function
'---------------------------------------------------------------------------------------------------
End If
'---------------------------------------------------------------------------------------------------
ReDim TEMP_MATRIX(0 To NROWS, 1 To 9)
'---------------------------------------------------------------------------------------------------
TEMP_MATRIX(0, 1) = "DATE"
TEMP_MATRIX(0, 2) = "OPEN"
TEMP_MATRIX(0, 3) = "CLOSE"
TEMP_MATRIX(0, 4) = "OPEN(t+1)"
TEMP_MATRIX(0, 5) = "Y-FIT"
TEMP_MATRIX(0, 6) = "RMS-ERROR"
TEMP_MATRIX(0, 7) = "WEIGHT"
TEMP_MATRIX(0, 8) = "LOWER BOUND"
TEMP_MATRIX(0, 9) = "UPPER BOUND"

RMS_ERROR = 0
MAX_ERROR = -2 ^ 52
AVG_ERROR = 0
WGT_ERROR = 0
TEMP_SUM = 0

For i = 1 To NROWS - 1
    TEMP_MATRIX(i, 1) = DATA_MATRIX(i, 1)
    TEMP_MATRIX(i, 2) = DATA_MATRIX(i, 2)
    TEMP_MATRIX(i, 3) = DATA_MATRIX(i, 3)
    TEMP_MATRIX(i, 4) = DATA_MATRIX(i + 1, 2)
    TEMP_MATRIX(i, 5) = ASSET_OPEN_PRICE_FORMULA_FUNC(TEMP_MATRIX(i, 3), PARAM_VECTOR, 1)
    
    TEMP_MATRIX(i, 6) = Abs(TEMP_MATRIX(i, 4) - TEMP_MATRIX(i, 5))
    If TEMP_MATRIX(i, 6) > MAX_ERROR Then: MAX_ERROR = TEMP_MATRIX(i, 6)
    RMS_ERROR = RMS_ERROR + TEMP_MATRIX(i, 6) ^ 2
    AVG_ERROR = AVG_ERROR + TEMP_MATRIX(i, 6)
    
    TEMP_MATRIX(i, 7) = WEIGHT_VAL ^ (NROWS - i) 'Weight
    TEMP_SUM = TEMP_SUM + TEMP_MATRIX(i, 7)
    WGT_ERROR = WGT_ERROR + TEMP_MATRIX(i, 6) * TEMP_MATRIX(i, 7)
    
    TEMP_MATRIX(i, 8) = ASSET_OPEN_PRICE_FORMULA_FUNC(TEMP_MATRIX(i, 3), _
                        PARAM_VECTOR, 1 - CONFIDENCE_VAL)
    
    TEMP_MATRIX(i, 9) = ASSET_OPEN_PRICE_FORMULA_FUNC(TEMP_MATRIX(i, 3), _
                        PARAM_VECTOR, 1 + CONFIDENCE_VAL)
Next i
i = NROWS
TEMP_MATRIX(i, 1) = DATA_MATRIX(i, 1)
TEMP_MATRIX(i, 2) = DATA_MATRIX(i, 2)
TEMP_MATRIX(i, 3) = DATA_MATRIX(i, 3)

Select Case OUTPUT
Case 1
    ASSET_OPEN_PRICE_FITNESS_FUNC = TEMP_MATRIX
Case Else
    RMS_ERROR = (RMS_ERROR / NROWS) ^ 0.5
    AVG_ERROR = AVG_ERROR / NROWS
    WGT_ERROR = WGT_ERROR / TEMP_SUM
    
    ASSET_OPEN_PRICE_FITNESS_FUNC = Array(RMS_ERROR, MAX_ERROR, AVG_ERROR, WGT_ERROR)
End Select

Exit Function
ERROR_LABEL:
ASSET_OPEN_PRICE_FITNESS_FUNC = Err.number
End Function

Function ASSET_OPEN_PRICE_OPTIMIZER_FUNC(ByRef XDATA_RNG As Variant, _
ByRef YDATA_RNG As Variant, _
ByRef PARAM_RNG As Variant, _
Optional ByVal WEIGHT_VAL As Double = 0.9, _
Optional ByVal OPTIM_TYPE As Integer = 0, _
Optional ByVal ERROR_TYPE As Integer = 0, _
Optional ByVal OUTPUT As Integer = 0)

Dim XDATA_VECTOR As Variant
Dim YDATA_VECTOR As Variant
Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL

XDATA_VECTOR = XDATA_RNG
If UBound(XDATA_VECTOR, 1) = 1 Then
    XDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(XDATA_VECTOR)
End If

YDATA_VECTOR = YDATA_RNG
If UBound(YDATA_VECTOR, 1) = 1 Then
    YDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(YDATA_VECTOR)
End If
If UBound(XDATA_VECTOR, 1) <> UBound(YDATA_VECTOR, 1) Then: GoTo ERROR_LABEL

PUB_XDATA_VECTOR = XDATA_VECTOR
PUB_YDATA_VECTOR = YDATA_VECTOR
PUB_WEIGHT_VAL = WEIGHT_VAL
PUB_ERROR_TYPE = ERROR_TYPE

PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then
    PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)
End If

'------------------------------------------------------------------------------------
If OUTPUT = 0 Then 'Optimal Parameters
'------------------------------------------------------------------------------------
    If OPTIM_TYPE = 0 Then
        ASSET_OPEN_PRICE_OPTIMIZER_FUNC = NELDER_MEAD_OPTIMIZATION2_FUNC( _
                                  "ASSET_OPEN_PRICE_ERROR_FUNC", _
                                  YDATA_VECTOR, PARAM_VECTOR)
    Else
        ASSET_OPEN_PRICE_OPTIMIZER_FUNC = NELDER_MEAD_OPTIMIZATION3_FUNC( _
                                  "ASSET_OPEN_PRICE_ERROR_FUNC", _
                                  PARAM_VECTOR, 1000, 10 ^ -10)
    End If
'------------------------------------------------------------------------------------
Else 'Error Val
'------------------------------------------------------------------------------------
    ASSET_OPEN_PRICE_OPTIMIZER_FUNC = ASSET_OPEN_PRICE_ERROR_FUNC(YDATA_VECTOR, PARAM_VECTOR)
'------------------------------------------------------------------------------------
End If
'------------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
'-----------------------------------------------------------------------------------------------------------
ASSET_OPEN_PRICE_OPTIMIZER_FUNC = PUB_EPSILON
'-----------------------------------------------------------------------------------------------------------
End Function
'-----------------------------------------------------------------------------------------------------------

Function ASSET_OPEN_PRICE_ERROR_FUNC(Optional ByRef YDATA_VECTOR As Variant, _
Optional ByVal PARAM_VECTOR As Variant)

Dim i As Long
Dim NROWS As Long

Dim X_VAL As Double
Dim Y_VAL As Double
Dim TEMP_SUM As Double

Dim YFIT_VAL As Double
Dim WEIGHT_VAL As Double

Dim RMS_ERROR As Double 'Error
Dim MAX_ERROR As Double 'Max Error
Dim WGT_ERROR As Double 'd Error

On Error GoTo ERROR_LABEL

NROWS = UBound(PUB_XDATA_VECTOR, 1)

'----------------------------------------------------------------------------------
Select Case PUB_ERROR_TYPE
'----------------------------------------------------------------------------------
Case 0 'min rms error
'----------------------------------------------------------------------------------
    RMS_ERROR = 0
    For i = 1 To NROWS
        GoSub ESTIMATE_LINE
        RMS_ERROR = RMS_ERROR + (Abs(Y_VAL - YFIT_VAL)) ^ 2
    Next i
    
    ASSET_OPEN_PRICE_ERROR_FUNC = (RMS_ERROR / NROWS) ^ 0.5
'----------------------------------------------------------------------------------
Case 1 'min max error
'----------------------------------------------------------------------------------
    MAX_ERROR = -2 ^ 52
    For i = 1 To NROWS
        GoSub ESTIMATE_LINE
        RMS_ERROR = Abs(Y_VAL - YFIT_VAL)
        If RMS_ERROR > MAX_ERROR Then: MAX_ERROR = RMS_ERROR
    Next i
    ASSET_OPEN_PRICE_ERROR_FUNC = MAX_ERROR
'----------------------------------------------------------------------------------
Case 2 'min avg error
'----------------------------------------------------------------------------------
    TEMP_SUM = 0
    For i = 1 To NROWS
        GoSub ESTIMATE_LINE
        RMS_ERROR = Abs(Y_VAL - YFIT_VAL)
        TEMP_SUM = TEMP_SUM + RMS_ERROR
    Next i
    ASSET_OPEN_PRICE_ERROR_FUNC = TEMP_SUM / NROWS

'----------------------------------------------------------------------------------
Case Else 'min wgt error
'----------------------------------------------------------------------------------
    'PARAM_VECTOR(4, 1) '0.9 --> Weight
    TEMP_SUM = 0
    For i = 1 To NROWS
        GoSub ESTIMATE_LINE
        RMS_ERROR = Abs(Y_VAL - YFIT_VAL)
        WEIGHT_VAL = PUB_WEIGHT_VAL ^ (NROWS - i)
        TEMP_SUM = TEMP_SUM + WEIGHT_VAL
        WGT_ERROR = WGT_ERROR + RMS_ERROR * WEIGHT_VAL
    Next i
    WGT_ERROR = WGT_ERROR / TEMP_SUM
    ASSET_OPEN_PRICE_ERROR_FUNC = WGT_ERROR
'----------------------------------------------------------------------------------
End Select
'----------------------------------------------------------------------------------

'--------------------------------------------------------------------
Exit Function
'--------------------------------------------------------------------
ESTIMATE_LINE:
'--------------------------------------------------------------------
    X_VAL = PUB_XDATA_VECTOR(i, 1)
    Y_VAL = PUB_YDATA_VECTOR(i, 1)
    YFIT_VAL = ASSET_OPEN_PRICE_FORMULA_FUNC(X_VAL, PARAM_VECTOR, 1)
'--------------------------------------------------------------------
Return
'--------------------------------------------------------------------
ERROR_LABEL:
ASSET_OPEN_PRICE_ERROR_FUNC = PUB_EPSILON
End Function

Function ASSET_OPEN_PRICE_FORMULA_FUNC(ByVal X_VAL As Double, _
ByVal PARAM_RNG As Variant, _
Optional ByVal CONFIDENCE_VAL As Double = 1)

Dim i As Long
Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL

PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then
    PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)
End If

For i = LBound(PARAM_VECTOR, 1) To UBound(PARAM_VECTOR, 1)
    PARAM_VECTOR(i, 1) = PARAM_VECTOR(i, 1) * CONFIDENCE_VAL
Next i

Select Case PUB_FORMULA_TYPE
Case 0
    ASSET_OPEN_PRICE_FORMULA_FUNC = PARAM_VECTOR(1, 1) + PARAM_VECTOR(2, 1) * _
                             X_VAL ^ 2 + PARAM_VECTOR(3, 1) * _
                             X_VAL ^ 4
Case Else
    ASSET_OPEN_PRICE_FORMULA_FUNC = -PARAM_VECTOR(1, 1) + PARAM_VECTOR(2, 1) * X_VAL - _
                            PARAM_VECTOR(3, 1) * X_VAL ^ 2
End Select

Exit Function
ERROR_LABEL:
ASSET_OPEN_PRICE_FORMULA_FUNC = Err.number
End Function
