Attribute VB_Name = "FINAN_PORT_MOMENTS_BIVAR_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_BIVAR_STATS_FUNC
'DESCRIPTION   : Portfolio Statistics
'LIBRARY       : PORTFOLIO
'GROUP         : MOMENTS_BIVAR
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************

'PORT_MOMENTS_TABLE_FUNC
Function PORT_BIVAR_STATS_FUNC(ByVal MEAN1_VAL As Double, _
ByVal MEAN2_VAL As Double, _
ByVal SIGMA1_VAL As Double, _
ByVal SIGMA2_VAL As Double, _
ByVal COVAR_VAL As Double, _
ByVal WEIGHT_VAL As Double, _
ByVal MIN_WEIGHT_VAL As Double, _
ByVal MAX_WEIGHT_VAL As Double, _
ByVal DELTA_WEIGHT_VAL As Double)

Dim i As Long
Dim NSIZE As Long
Dim TEMP_SUM As Double 'Weights
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL
    
    NSIZE = (MAX_WEIGHT_VAL - MIN_WEIGHT_VAL) / DELTA_WEIGHT_VAL + 1
    
    ReDim TEMP_MATRIX(1 To 4, 1 To NSIZE + 2)
    
    TEMP_MATRIX(1, 1) = ("PORT MEAN")
    TEMP_MATRIX(2, 1) = ("PORT SIGMA")
    TEMP_MATRIX(3, 1) = ("PORT CORREL")
    TEMP_MATRIX(4, 1) = ("PORT WEIGHTS")
    TEMP_SUM = WEIGHT_VAL
        
    TEMP_MATRIX(1, 2) = PORT_BIVAR_MEAN_FUNC(TEMP_SUM, MEAN1_VAL, MEAN2_VAL)
    TEMP_MATRIX(2, 2) = PORT_BIVAR_VOLATILITY_FUNC(TEMP_SUM, SIGMA1_VAL, SIGMA2_VAL, COVAR_VAL)
    TEMP_MATRIX(3, 2) = PORT_BIVAR_CORREL_FUNC(COVAR_VAL, SIGMA1_VAL, SIGMA2_VAL)
    TEMP_MATRIX(4, 2) = TEMP_SUM
    
    For i = 1 To NSIZE
        TEMP_SUM = TEMP_SUM + DELTA_WEIGHT_VAL
        TEMP_MATRIX(1, i + 2) = PORT_BIVAR_MEAN_FUNC(TEMP_SUM, MEAN1_VAL, MEAN2_VAL)
        TEMP_MATRIX(2, i + 2) = PORT_BIVAR_VOLATILITY_FUNC(TEMP_SUM, SIGMA1_VAL, SIGMA2_VAL, COVAR_VAL)
        TEMP_MATRIX(3, i + 2) = PORT_BIVAR_CORREL_FUNC(COVAR_VAL, SIGMA1_VAL, SIGMA2_VAL)
        TEMP_MATRIX(4, i + 2) = TEMP_SUM
    Next i
    
    PORT_BIVAR_STATS_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
PORT_BIVAR_STATS_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_BIVAR_MEAN_FUNC
'DESCRIPTION   : Returns the portfolio mean
'LIBRARY       : PORTFOLIO
'GROUP         : MOMENTS_BIVAR
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************


Function PORT_BIVAR_MEAN_FUNC(ByVal WEIGHT_VAL As Double, _
ByVal MEAN1_VAL As Double, _
ByVal MEAN2_VAL As Double)

On Error GoTo ERROR_LABEL
    
PORT_BIVAR_MEAN_FUNC = WEIGHT_VAL * MEAN1_VAL + (1 - WEIGHT_VAL) * MEAN2_VAL

Exit Function
ERROR_LABEL:
PORT_BIVAR_MEAN_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_BIVAR_VOLATILITY_FUNC
'DESCRIPTION   : Returns the portfolio standard deviation
'LIBRARY       : PORTFOLIO
'GROUP         : MOMENTS_BIVAR
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************

Function PORT_BIVAR_VOLATILITY_FUNC(ByVal WEIGHT_VAL As Double, _
ByVal SIGMA1_VAL As Double, _
ByVal SIGMA2_VAL As Double, _
ByVal COVAR_VAL As Double)

Dim CORREL As Double

On Error GoTo ERROR_LABEL
    
CORREL = PORT_BIVAR_CORREL_FUNC(COVAR_VAL, SIGMA1_VAL, SIGMA2_VAL)
PORT_BIVAR_VOLATILITY_FUNC = PORT_BIVAR_VARIANCE_FUNC(WEIGHT_VAL, SIGMA1_VAL, SIGMA2_VAL, COVAR_VAL) ^ 0.5

Exit Function
ERROR_LABEL:
PORT_BIVAR_VOLATILITY_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_BIVAR_VARIANCE_FUNC
'DESCRIPTION   : Returns the portfolio variance
'LIBRARY       : PORTFOLIO
'GROUP         : MOMENTS_BIVAR
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************


Function PORT_BIVAR_VARIANCE_FUNC(ByVal WEIGHT_VAL As Double, _
ByVal SIGMA1_VAL As Double, _
ByVal SIGMA2_VAL As Double, _
ByVal COVAR_VAL As Double)

Dim CORREL As Double

On Error GoTo ERROR_LABEL
    CORREL = PORT_BIVAR_CORREL_FUNC(COVAR_VAL, SIGMA1_VAL, SIGMA2_VAL)
    
    PORT_BIVAR_VARIANCE_FUNC = _
    WEIGHT_VAL ^ 2 * (SIGMA1_VAL) ^ 2 + (1 - WEIGHT_VAL) ^ 2 * (SIGMA2_VAL) ^ 2 + 2 * WEIGHT_VAL * (1 - WEIGHT_VAL) * CORREL * (SIGMA1_VAL) * (SIGMA2_VAL)

Exit Function
ERROR_LABEL:
PORT_BIVAR_VARIANCE_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_BIVAR_CORREL_FUNC
'DESCRIPTION   : Returns the portfolio correlation
'LIBRARY       : PORTFOLIO
'GROUP         : MOMENTS_BIVAR
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************

Function PORT_BIVAR_CORREL_FUNC(ByVal COVAR_VAL As Double, _
ByVal SIGMA1_VAL As Double, _
ByVal SIGMA2_VAL As Double)

On Error GoTo ERROR_LABEL

PORT_BIVAR_CORREL_FUNC = COVAR_VAL / (SIGMA1_VAL * SIGMA2_VAL)

Exit Function
ERROR_LABEL:
PORT_BIVAR_CORREL_FUNC = Err.number
End Function
