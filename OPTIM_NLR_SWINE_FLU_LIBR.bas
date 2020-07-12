Attribute VB_Name = "OPTIM_NLR_SWINE_FLU_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

Private Const PUB_EPSILON As Double = 2 ^ 52 '1E-100
'Private Const PUB_FORMULA_TYPE As Integer = 1
'another option is to use  K/(1+A*EXP(-b*x))
Private PUB_WEIGHT_VAL As Double
Private PUB_YDATA_VECTOR As Variant
Private PUB_OBJ_TYPE As Integer
Private PUB_OPTIM_TYPE As Integer

'////////////////////////////////////////////////////////////////////////////////////////////////////
'http://www.gummy-stuff.org/swine-flu.htm
'http://mathworld.wolfram.com/LeastSquaresFittingExponential.html
'////////////////////////////////////////////////////////////////////////////////////////////////////

Function SWINE_FLU_CASES_FITNESS_FUNC(ByRef YDATA_RNG As Variant, _
ByRef PARAM_RNG As Variant, _
Optional ByVal OPTIM_TYPE As Integer = 0, _
Optional ByVal OBJ_TYPE As Integer = 0, _
Optional ByVal CONFIDENCE_VAL As Double = 0.01, _
Optional ByVal WEIGHT_VAL As Double = 0.9, _
Optional ByVal OUTPUT As Integer = 0)

Dim i As Long
Dim NROWS As Long

Dim RMS_ERROR As Double 'Error
Dim MAX_ERROR As Double 'Max Error
Dim AVG_ERROR As Double 'AVG Error
Dim WGT_ERROR As Double 'd Error
Dim TEMP_SUM As Double 'weight factor

Dim TEMP_MATRIX As Variant
Dim YDATA_VECTOR As Variant
Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL

YDATA_VECTOR = YDATA_RNG
If UBound(YDATA_VECTOR, 1) = 1 Then
    YDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(YDATA_VECTOR)
End If
NROWS = UBound(YDATA_VECTOR, 1)

'PARAM_VECTOR(1, 1) --> 0.0319999977946281
'PARAM_VECTOR(2, 1) --> 11.5280017852783
'PARAM_VECTOR(3, 1) --> 0.460000067949295
'PARAM_VECTOR(4, 1) --> 1.03600001335144

PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then
    PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)
End If

'---------------------------------------------------------------------------------------------------
If OUTPUT = 0 Then
'---------------------------------------------------------------------------------------------------
    SWINE_FLU_CASES_FITNESS_FUNC = SWINE_FLU_CASES_OPTIMIZER_FUNC(YDATA_VECTOR, _
                                   PARAM_VECTOR, OPTIM_TYPE, OBJ_TYPE, WEIGHT_VAL, 0)
    Exit Function
'---------------------------------------------------------------------------------------------------
End If
'---------------------------------------------------------------------------------------------------

ReDim TEMP_MATRIX(0 To NROWS, 1 To 9)
TEMP_MATRIX(0, 1) = "INDEX"
TEMP_MATRIX(0, 2) = "CASES"
TEMP_MATRIX(0, 3) = "LOG(CASES)"
TEMP_MATRIX(0, 4) = "LOG(FIT)"
TEMP_MATRIX(0, 5) = "RMS-ERROR"
TEMP_MATRIX(0, 6) = "Y-FIT"
TEMP_MATRIX(0, 7) = "WEIGHT"
TEMP_MATRIX(0, 8) = "LOWER BOUND"
TEMP_MATRIX(0, 9) = "UPPER BOUND"

RMS_ERROR = 0
MAX_ERROR = -2 ^ 52
AVG_ERROR = 0
WGT_ERROR = 0
TEMP_SUM = 0

For i = 1 To NROWS
    TEMP_MATRIX(i, 1) = i
    TEMP_MATRIX(i, 2) = YDATA_VECTOR(i, 1)
    TEMP_MATRIX(i, 3) = Log(TEMP_MATRIX(i, 2))
    TEMP_MATRIX(i, 4) = SWINE_FLU_CASES_FORMULA_FUNC(CDbl(i), PARAM_VECTOR, 1, 0)
    
    TEMP_MATRIX(i, 5) = Abs(TEMP_MATRIX(i, 3) - TEMP_MATRIX(i, 4))
    If TEMP_MATRIX(i, 5) > MAX_ERROR Then: MAX_ERROR = TEMP_MATRIX(i, 5)
    
    RMS_ERROR = RMS_ERROR + TEMP_MATRIX(i, 5) ^ 2
    AVG_ERROR = AVG_ERROR + TEMP_MATRIX(i, 5)
    
    TEMP_MATRIX(i, 6) = Exp(TEMP_MATRIX(i, 4))
    TEMP_MATRIX(i, 7) = WEIGHT_VAL ^ (NROWS - i)
    TEMP_SUM = TEMP_SUM + TEMP_MATRIX(i, 7)
    WGT_ERROR = WGT_ERROR + TEMP_MATRIX(i, 5) * TEMP_MATRIX(i, 7)
    
    TEMP_MATRIX(i, 8) = SWINE_FLU_CASES_FORMULA_FUNC(CDbl(i), PARAM_VECTOR, CONFIDENCE_VAL, 1)
    TEMP_MATRIX(i, 8) = Exp(TEMP_MATRIX(i, 8))
    
    TEMP_MATRIX(i, 9) = SWINE_FLU_CASES_FORMULA_FUNC(CDbl(i), PARAM_VECTOR, CONFIDENCE_VAL, 2)
    TEMP_MATRIX(i, 9) = Exp(TEMP_MATRIX(i, 9))
    
Next i
RMS_ERROR = (RMS_ERROR / NROWS) ^ 0.5
AVG_ERROR = AVG_ERROR / NROWS
WGT_ERROR = WGT_ERROR / TEMP_SUM

Select Case OUTPUT
Case 1
    SWINE_FLU_CASES_FITNESS_FUNC = TEMP_MATRIX
Case Else
    SWINE_FLU_CASES_FITNESS_FUNC = Array(RMS_ERROR, MAX_ERROR, AVG_ERROR, WGT_ERROR)
End Select

'--------------------------------------------------------------------
Exit Function
'--------------------------------------------------------------------
ERROR_LABEL:
SWINE_FLU_CASES_FITNESS_FUNC = Err.number
End Function

Function SWINE_FLU_CASES_OPTIMIZER_FUNC(ByRef YDATA_RNG As Variant, _
ByRef PARAM_RNG As Variant, _
Optional ByVal OPTIM_TYPE As Integer = 0, _
Optional ByVal OBJ_TYPE As Integer = 0, _
Optional ByVal WEIGHT_VAL As Double = 0.9, _
Optional ByVal OUTPUT As Integer = 0)

Dim YDATA_VECTOR As Variant
Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL

YDATA_VECTOR = YDATA_RNG
If UBound(YDATA_VECTOR, 1) = 1 Then
    YDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(YDATA_VECTOR)
End If

PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then
    PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)
End If
PUB_OBJ_TYPE = OBJ_TYPE
PUB_OPTIM_TYPE = OPTIM_TYPE
PUB_WEIGHT_VAL = WEIGHT_VAL
PUB_YDATA_VECTOR = YDATA_VECTOR

If OUTPUT = 0 Then 'Optimal Parameters
    If OPTIM_TYPE = 0 Then
        SWINE_FLU_CASES_OPTIMIZER_FUNC = NELDER_MEAD_OPTIMIZATION2_FUNC( _
                                  "SWINE_FLU_CASES_ERROR_FUNC", _
                                  YDATA_VECTOR, PARAM_VECTOR)
    Else
        SWINE_FLU_CASES_OPTIMIZER_FUNC = NELDER_MEAD_OPTIMIZATION3_FUNC( _
                                  "SWINE_FLU_CASES_ERROR_FUNC", _
                                  PARAM_VECTOR, 1000, 10 ^ -10)
    End If

Else 'Error Val
    SWINE_FLU_CASES_OPTIMIZER_FUNC = SWINE_FLU_CASES_ERROR_FUNC(YDATA_VECTOR, PARAM_VECTOR)
End If

Exit Function
ERROR_LABEL:
SWINE_FLU_CASES_OPTIMIZER_FUNC = PUB_EPSILON
End Function

Function SWINE_FLU_CASES_ERROR_FUNC(Optional ByVal YDATA_VECTOR As Variant, _
Optional ByVal PARAM_VECTOR As Variant)

Dim i As Long
Dim NROWS As Long

Dim X_VAL As Double
Dim Y_VAL As Double
Dim YFIT_VAL As Double
Dim WEIGHT_VAL As Double
Dim TEMP_SUM As Double 'weight factor

Dim RMS_ERROR As Double 'Error
Dim MAX_ERROR As Double 'Max Error
Dim WGT_ERROR As Double 'd Error

On Error GoTo ERROR_LABEL

If PUB_OPTIM_TYPE <> 0 Then: PARAM_VECTOR = YDATA_VECTOR
NROWS = UBound(PUB_YDATA_VECTOR, 1)

'----------------------------------------------------------------------------------
Select Case PUB_OBJ_TYPE
'----------------------------------------------------------------------------------
Case 0 'min rms error
'----------------------------------------------------------------------------------
    RMS_ERROR = 0
    For i = 1 To NROWS
        GoSub ESTIMATE_LINE
        RMS_ERROR = RMS_ERROR + (Abs(Y_VAL - YFIT_VAL)) ^ 2
    Next i
    
    SWINE_FLU_CASES_ERROR_FUNC = (RMS_ERROR / NROWS) ^ 0.5
'----------------------------------------------------------------------------------
Case 1 'min max error
'----------------------------------------------------------------------------------
    MAX_ERROR = -2 ^ 52
    For i = 1 To NROWS
        GoSub ESTIMATE_LINE
        RMS_ERROR = Abs(Y_VAL - YFIT_VAL)
        If RMS_ERROR > MAX_ERROR Then: MAX_ERROR = RMS_ERROR
    Next i
    SWINE_FLU_CASES_ERROR_FUNC = MAX_ERROR
'----------------------------------------------------------------------------------
Case 2 'min avg error
'----------------------------------------------------------------------------------
    TEMP_SUM = 0
    For i = 1 To NROWS
        GoSub ESTIMATE_LINE
        RMS_ERROR = Abs(Y_VAL - YFIT_VAL)
        TEMP_SUM = TEMP_SUM + RMS_ERROR
    Next i
    SWINE_FLU_CASES_ERROR_FUNC = TEMP_SUM / NROWS
'----------------------------------------------------------------------------------
Case Else 'min wgt error
'----------------------------------------------------------------------------------
    TEMP_SUM = 0
    For i = 1 To NROWS
        GoSub ESTIMATE_LINE
        RMS_ERROR = Abs(Y_VAL - YFIT_VAL)
        WEIGHT_VAL = PUB_WEIGHT_VAL ^ (NROWS - i)
        TEMP_SUM = TEMP_SUM + WEIGHT_VAL
        WGT_ERROR = WGT_ERROR + RMS_ERROR * WEIGHT_VAL
    Next i
    WGT_ERROR = WGT_ERROR / TEMP_SUM
    SWINE_FLU_CASES_ERROR_FUNC = WGT_ERROR
'----------------------------------------------------------------------------------
End Select
'----------------------------------------------------------------------------------

'--------------------------------------------------------------------
Exit Function
'--------------------------------------------------------------------
ESTIMATE_LINE:
'--------------------------------------------------------------------
    X_VAL = i
    Y_VAL = Log(PUB_YDATA_VECTOR(i, 1))
    YFIT_VAL = SWINE_FLU_CASES_FORMULA_FUNC(X_VAL, PARAM_VECTOR, 1, 0)
'--------------------------------------------------------------------
Return
'--------------------------------------------------------------------
ERROR_LABEL:
'--------------------------------------------------------------------
SWINE_FLU_CASES_ERROR_FUNC = PUB_EPSILON
End Function


Function SWINE_FLU_CASES_FORMULA_FUNC(ByVal X_VAL As Double, _
ByVal PARAM_VECTOR As Variant, _
Optional ByVal CONFIDENCE_VAL As Double = 1, _
Optional ByVal VERSION As Integer = 0)

'rr --> PARAM_VECTOR(1, 1) --> 0.0319999977946281
'K --> PARAM_VECTOR(2, 1) --> 11.5280017852783
'A --> PARAM_VECTOR(3, 1) --> 0.460000067949295
'b --> PARAM_VECTOR(4, 1) --> 1.03600001335144

Dim i As Long
On Error GoTo ERROR_LABEL

'---------------------------------------------------------------------------------------
Select Case VERSION 'Fitted to : log(y) = K*(1-A*EXP(-rr*x))^b
'---------------------------------------------------------------------------------------
Case 0 'Base
    For i = LBound(PARAM_VECTOR, 1) To UBound(PARAM_VECTOR, 1)
        PARAM_VECTOR(i, 1) = PARAM_VECTOR(i, 1) * CONFIDENCE_VAL
    Next i
    
    '=K/(1+A*EXP(-b*x))^rr
    SWINE_FLU_CASES_FORMULA_FUNC = PARAM_VECTOR(2, 1) * (1 - _
                                   PARAM_VECTOR(3, 1) * Exp(-PARAM_VECTOR(1, 1) * _
                                   X_VAL)) ^ PARAM_VECTOR(4, 1)
'---------------------------------------------------------------------------------------
Case 1 'Lower Bound
'---------------------------------------------------------------------------------------
    '=K/(1+A*EXP(-b*x))^rr
    SWINE_FLU_CASES_FORMULA_FUNC = (1 - CONFIDENCE_VAL) * PARAM_VECTOR(2, 1) * (1 - _
                                   (1 + CONFIDENCE_VAL) * PARAM_VECTOR(3, 1) * Exp(-(1 + _
                                   CONFIDENCE_VAL) * PARAM_VECTOR(1, 1) * X_VAL)) ^ PARAM_VECTOR(4, 1)
'---------------------------------------------------------------------------------------
Case Else 'Upper Bound
'---------------------------------------------------------------------------------------
    '=K/(1+A*EXP(-b*x))^rr
    SWINE_FLU_CASES_FORMULA_FUNC = (1 + CONFIDENCE_VAL) * PARAM_VECTOR(2, 1) * (1 - _
                                   (1 - CONFIDENCE_VAL) * PARAM_VECTOR(3, 1) * Exp(-(1 - _
                                   CONFIDENCE_VAL) * PARAM_VECTOR(1, 1) * X_VAL)) ^ PARAM_VECTOR(4, 1)
'---------------------------------------------------------------------------------------
End Select
'---------------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
SWINE_FLU_CASES_FORMULA_FUNC = Err.number
End Function

