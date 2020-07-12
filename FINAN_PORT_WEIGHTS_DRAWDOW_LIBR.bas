Attribute VB_Name = "FINAN_PORT_WEIGHTS_DRAWDOW_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

Private PUB_DATA_MATRIX As Variant
Private PUB_MIN_EXPOS_VAL As Double
Private PUB_MAX_EXPOS_VAL As Double
Private Const PUB_EPSILON As Double = 2 ^ 52


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : PORT_MAX_DRAWDOWN_WEIGHTS_OPTIMIZER_FUNC

'DESCRIPTION   : An optimizer based on an open source genetic algorithm
'(pikaia) maximizing Return ^ 3 / Max Drawdown ^ 2 under 135/35 restrictions.

'LIBRARY       : PORTFOLIO
'GROUP         : WEIGHTS_DRAWDOWN
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'**********************************************************************************
'**********************************************************************************

Function PORT_MAX_DRAWDOWN_WEIGHTS_OPTIMIZER_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal DATA_TYPE As Integer = 0, _
Optional ByVal LOG_SCALE As Integer = 1, _
Optional ByRef WEIGHTS_RNG As Variant, _
Optional ByRef LOWER_RNG As Variant, _
Optional ByRef UPPER_RNG As Variant, _
Optional ByVal MIN_EXPOS_VAL As Double = 0.8, _
Optional ByVal MAX_EXPOS_VAL As Double = 1, _
Optional ByVal TRACE_FLAG As Boolean = False)

'MIN_EXPOS_VAL = Minimum Exposure
'MAX_EXPOS_VAL = Maximum Exposure

Dim i As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim YTEMP_VAL As Double

Dim CONST_BOX As Variant
Dim RETURNS_MATRIX As Variant
Dim WEIGHTS_VECTOR As Variant
Dim LOWER_VECTOR As Variant 'Minimum Exposure
Dim UPPER_VECTOR As Variant 'Maximum Exposure

Dim ERROR_STR As String

On Error GoTo ERROR_LABEL

PUB_DATA_MATRIX = 0
PUB_MIN_EXPOS_VAL = 0
PUB_MAX_EXPOS_VAL = 0

RETURNS_MATRIX = DATA_RNG
If DATA_TYPE <> 0 Then: RETURNS_MATRIX = MATRIX_PERCENT_FUNC(RETURNS_MATRIX, LOG_SCALE)
NROWS = UBound(RETURNS_MATRIX, 1)
NCOLUMNS = UBound(RETURNS_MATRIX, 2)

PUB_DATA_MATRIX = RETURNS_MATRIX
PUB_MIN_EXPOS_VAL = MIN_EXPOS_VAL
PUB_MAX_EXPOS_VAL = MAX_EXPOS_VAL

'----------------------------------------------------------------------------------------
If IsArray(WEIGHTS_RNG) = False Then
'----------------------------------------------------------------------------------------
    
    LOWER_VECTOR = LOWER_RNG
    If UBound(LOWER_VECTOR, 1) = 1 Then
        LOWER_VECTOR = MATRIX_TRANSPOSE_FUNC(LOWER_VECTOR)
    End If
    
    UPPER_VECTOR = UPPER_RNG
    If UBound(UPPER_VECTOR, 1) = 1 Then
        UPPER_VECTOR = MATRIX_TRANSPOSE_FUNC(UPPER_VECTOR)
    End If
    ReDim CONST_BOX(1 To 2, 1 To NCOLUMNS)
    For i = 1 To NCOLUMNS
        CONST_BOX(1, i) = LOWER_VECTOR(i, 1)
        CONST_BOX(2, i) = UPPER_VECTOR(i, 1)
    Next i
    WEIGHTS_VECTOR = PIKAIA_OPTIMIZATION_FUNC("PORT_MAX_DRAWDOWN_WEIGHTS_OBJ_FUNC", CONST_BOX, TRACE_FLAG, ERROR_STR, 123456, 100, 50, 5, 0.85, 2, 0.005, 0.0005, 0.25, 1, 1, 1, 0)
    If ERROR_STR <> "" Then: GoTo ERROR_LABEL
    PORT_MAX_DRAWDOWN_WEIGHTS_OPTIMIZER_FUNC = WEIGHTS_VECTOR
 '----------------------------------------------------------------------------------------
Else
'----------------------------------------------------------------------------------------
    WEIGHTS_VECTOR = WEIGHTS_RNG
    If UBound(WEIGHTS_VECTOR, 1) = 1 Then: _
        WEIGHTS_VECTOR = MATRIX_TRANSPOSE_FUNC(WEIGHTS_VECTOR)
    YTEMP_VAL = PORT_MAX_DRAWDOWN_WEIGHTS_OBJ_FUNC(WEIGHTS_VECTOR)
    
    PORT_MAX_DRAWDOWN_WEIGHTS_OPTIMIZER_FUNC = YTEMP_VAL
'----------------------------------------------------------------------------------------
End If
'----------------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
PORT_MAX_DRAWDOWN_WEIGHTS_OPTIMIZER_FUNC = Err.number
End Function


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : PORT_MAX_DRAWDOWN_WEIGHTS_OBJ_FUNC
'DESCRIPTION   : The Genetic Return^3-Drawdown^2 135/35
'Objective Function
'LIBRARY       : PORTFOLIO
'GROUP         : WEIGHTS_DRAWDOWN
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'**********************************************************************************
'**********************************************************************************

Private Function PORT_MAX_DRAWDOWN_WEIGHTS_OBJ_FUNC(ByRef WEIGHTS_RNG As Variant)

Dim i As Long
Dim j As Long

Dim NSIZE As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_SUM As Double
Dim TEMP_DIFF As Double

Dim MEAN_VAL As Double
Dim WEIGHTS_SUM As Double
Dim MAX_DRAWDOWN_VAL As Double

Dim YTEMP_VAL As Double
Dim XTEMP_VAL As Double

Dim MIN_EXPOS_VAL As Double
Dim MAX_EXPOS_VAL As Double

Dim RETURNS_MATRIX As Variant
Dim WEIGHTS_VECTOR As Variant

Dim OPTIMAL_VECTOR As Variant

Dim tolerance As Double
Dim epsilon As Double

On Error GoTo ERROR_LABEL

epsilon = 10 ^ -15
tolerance = 1E+100
RETURNS_MATRIX = PUB_DATA_MATRIX
NROWS = UBound(RETURNS_MATRIX, 1)
NCOLUMNS = UBound(RETURNS_MATRIX, 2)

WEIGHTS_VECTOR = WEIGHTS_RNG
If UBound(WEIGHTS_VECTOR, 1) = 1 Then
    WEIGHTS_VECTOR = MATRIX_TRANSPOSE_FUNC(WEIGHTS_VECTOR)
End If
NSIZE = UBound(WEIGHTS_VECTOR, 1)
If NSIZE <> NCOLUMNS Then: GoTo ERROR_LABEL

MIN_EXPOS_VAL = PUB_MIN_EXPOS_VAL
MAX_EXPOS_VAL = PUB_MAX_EXPOS_VAL

'-----------------------------------------------------------------------
ReDim OPTIMAL_VECTOR(1 To NROWS, 1 To 1)
'-----------------------------------------------------------------------
MEAN_VAL = 0
For i = 1 To NROWS
    TEMP_SUM = 0
    For j = 1 To NCOLUMNS
        TEMP_SUM = TEMP_SUM + RETURNS_MATRIX(i, j) * WEIGHTS_VECTOR(j, 1)
    Next j
    OPTIMAL_VECTOR(i, 1) = TEMP_SUM
    MEAN_VAL = MEAN_VAL + OPTIMAL_VECTOR(i, 1)
Next i
MEAN_VAL = MEAN_VAL / NROWS
'-----------------------------------------------------------------------
MAX_DRAWDOWN_VAL = 0
TEMP_SUM = 0
For i = 1 To NROWS - 1
    TEMP_DIFF = OPTIMAL_VECTOR(i + 1, 1) - OPTIMAL_VECTOR(i, 1)
    If TEMP_DIFF <= 0 Then
        TEMP_SUM = TEMP_SUM + TEMP_DIFF
    ElseIf MAX_DRAWDOWN_VAL > TEMP_SUM Then
        MAX_DRAWDOWN_VAL = TEMP_SUM
        TEMP_SUM = 0
    End If
Next i
If MAX_DRAWDOWN_VAL ^ 2 <= epsilon Then: MAX_DRAWDOWN_VAL = tolerance
'-----------------------------------------------------------------------

WEIGHTS_SUM = 0 'actual exposure
For j = 1 To NCOLUMNS: WEIGHTS_SUM = WEIGHTS_SUM + WEIGHTS_VECTOR(j, 1): Next j
'-----------------------------------------------------------------------

If WEIGHTS_SUM > MAX_EXPOS_VAL Then
    XTEMP_VAL = PUB_EPSILON
Else
    If WEIGHTS_SUM < MIN_EXPOS_VAL Then
        XTEMP_VAL = PUB_EPSILON
    Else
        XTEMP_VAL = 1
    End If
End If

YTEMP_VAL = (MEAN_VAL ^ 3 / MAX_DRAWDOWN_VAL ^ 2) / XTEMP_VAL

PORT_MAX_DRAWDOWN_WEIGHTS_OBJ_FUNC = YTEMP_VAL

Exit Function
ERROR_LABEL:
PORT_MAX_DRAWDOWN_WEIGHTS_OBJ_FUNC = 1 / PUB_EPSILON
End Function
