Attribute VB_Name = "STAT_REGRESSION_QUANT_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : QUANTILE_REGRESSION_FUNC

'DESCRIPTION   : http://www.econ.uiuc.edu/~roger/research/rq/QRJEP.pdf
'LIBRARY       : STATISTICAL
'GROUP         : QUANTILE
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function QUANTILE_REGRESSION_FUNC(ByRef XDATA_RNG As Variant, _
ByRef YDATA_RNG As Variant, _
Optional ByVal QUANTILE As Double = 0.1, _
Optional ByRef PARAM_RNG As Variant, _
Optional ByVal nLOOPS As Long = 1000, _
Optional ByVal tolerance As Double = 0.000000000000001, _
Optional ByVal OUTPUT As Integer = 0)

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_VECTOR As Variant
Dim TEMP_MATRIX As Variant

Dim PARAM_VECTOR As Variant

Dim YDATA_VECTOR As Variant
Dim XDATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

XDATA_MATRIX = XDATA_RNG
YDATA_VECTOR = YDATA_RNG
If UBound(YDATA_VECTOR, 1) = 1 Then
    YDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(YDATA_VECTOR)
End If
NROWS = UBound(YDATA_VECTOR, 1)
If UBound(XDATA_MATRIX, 1) <> NROWS Then: GoTo ERROR_LABEL
NCOLUMNS = UBound(XDATA_MATRIX, 2)

'----------------------------------------------------------------------------------
If IsArray(PARAM_RNG) = True Then
'----------------------------------------------------------------------------------
    PARAM_VECTOR = PARAM_RNG
    If UBound(PARAM_VECTOR, 1) = 1 Then
        PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)
    End If
    If NCOLUMNS <> UBound(PARAM_VECTOR, 1) Then: GoTo ERROR_LABEL
    
'----------------------------------------------------------------------------------
Else
'----------------------------------------------------------------------------------

    ReDim PARAM_VECTOR(1 To NCOLUMNS, 1 To 1)
    For i = 1 To NCOLUMNS
        PARAM_VECTOR(i, 1) = 1
    Next i
    
    PARAM_VECTOR = NELDER_MEAD_OPTIMIZATION4_FUNC("QUANTILE_OBJECTIVE_FUNC", _
                   XDATA_MATRIX, YDATA_VECTOR, QUANTILE, PARAM_VECTOR, 5000, 0.0000000001)
    
'    PARAM_VECTOR = NELDER_MEAD_OPTIMIZATION_FRAME_FUNC("QUANTILE_OBJECTIVE_FUNC", _
                   PARAM_VECTOR, "", True, 0, nLOOPS, tolerance)

'----------------------------------------------------------------------------------
End If
'----------------------------------------------------------------------------------
  
'----------------------------------------------------------------------------------
Select Case OUTPUT
'----------------------------------------------------------------------------------
    Case 0
'----------------------------------------------------------------------------------
        ReDim TEMP_MATRIX(0 To NROWS, 1 To 4)
        ReDim TEMP_VECTOR(1 To 1, 1 To NCOLUMNS)
        TEMP_MATRIX(0, 1) = "Y_HAT"
        TEMP_MATRIX(0, 2) = "{e}"
        TEMP_MATRIX(0, 3) = "{p}"
        TEMP_MATRIX(0, 4) = "{p*e}"
        
        For i = 1 To NROWS
            For j = 1 To NCOLUMNS
                TEMP_VECTOR(1, j) = XDATA_MATRIX(i, j)
            Next j
            TEMP_MATRIX(i, 1) = MMULT_FUNC(TEMP_VECTOR, PARAM_VECTOR, 70)(1, 1)
            TEMP_MATRIX(i, 2) = YDATA_VECTOR(i, 1) - TEMP_MATRIX(i, 1)
            TEMP_MATRIX(i, 3) = QUANTILE - IIf(TEMP_MATRIX(i, 2) < 0, 1, 0)
            TEMP_MATRIX(i, 4) = TEMP_MATRIX(i, 2) * TEMP_MATRIX(i, 3)
        Next i
            
        QUANTILE_REGRESSION_FUNC = TEMP_MATRIX
'----------------------------------------------------------------------------------
    Case 1
'----------------------------------------------------------------------------------
        ReDim TEMP_VECTOR(1 To 1, 1 To 1)
        TEMP_VECTOR(1, 1) = QUANTILE
        QUANTILE_REGRESSION_FUNC = QUANTILE_OBJECTIVE_FUNC(XDATA_MATRIX, YDATA_VECTOR, _
                                   TEMP_VECTOR, PARAM_VECTOR)
'----------------------------------------------------------------------------------
    Case Else
'----------------------------------------------------------------------------------
        QUANTILE_REGRESSION_FUNC = PARAM_VECTOR
'----------------------------------------------------------------------------------
End Select
'----------------------------------------------------------------------------------
    
Exit Function
ERROR_LABEL:
QUANTILE_REGRESSION_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : QUANTILE_OBJECTIVE_FUNC
'DESCRIPTION   : Quantile Regression Objective Function
'LIBRARY       : STATISTICAL
'GROUP         : QUANTILE
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************


Private Function QUANTILE_OBJECTIVE_FUNC(ByRef XDATA_RNG As Variant, _
ByRef YDATA_RNG As Variant, _
ByRef CONST_RNG As Variant, _
ByRef PARAM_RNG As Variant)

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim YHAT_VAL As Double
Dim MULT_VAL As Double
Dim TEMP_SUM As Double
Dim ERROR_VAL As Double

On Error GoTo ERROR_LABEL

NROWS = UBound(YDATA_RNG, 1)
NCOLUMNS = UBound(XDATA_RNG, 2)

TEMP_SUM = 0

For i = 1 To NROWS
    YHAT_VAL = 0
    For j = 1 To NCOLUMNS
        YHAT_VAL = YHAT_VAL + PARAM_RNG(j, 1) * XDATA_RNG(i, j)
    Next j
    ERROR_VAL = YDATA_RNG(i, 1) - YHAT_VAL
    If ERROR_VAL < 0 Then
        MULT_VAL = ERROR_VAL * (CONST_RNG(1, 1) - 1)
    Else
        MULT_VAL = ERROR_VAL * CONST_RNG(1, 1)
    End If
    TEMP_SUM = TEMP_SUM + MULT_VAL
Next i

QUANTILE_OBJECTIVE_FUNC = TEMP_SUM

Exit Function
ERROR_LABEL:
QUANTILE_OBJECTIVE_FUNC = Err.number
End Function
