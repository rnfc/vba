Attribute VB_Name = "FINAN_PORT_MOMENTS_TIMING_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_MARKET_TIMING_FUNC
'DESCRIPTION   : INDEX ; SQUARED; HENRIKSON/MERTON; UP & DOWN MODELS
'LIBRARY       : FINAN_PORT
'GROUP         : TIMING
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 03/16/2009
'************************************************************************************
'************************************************************************************

Function PORT_MARKET_TIMING_FUNC(ByRef XDATA_RNG As Variant, _
ByRef YDATA_RNG As Variant, _
Optional ByVal DATA_TYPE As Integer = 0, _
Optional ByVal LOG_SCALE As Integer = 0, _
Optional ByVal OUTPUT As Integer = 0)

Dim i As Long
Dim NROWS As Long

Dim XDATA_VECTOR As Variant
Dim YDATA_VECTOR As Variant

Dim XDATA_MATRIX As Variant

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

If DATA_TYPE <> 0 Then: XDATA_VECTOR = MATRIX_PERCENT_FUNC(XDATA_VECTOR, LOG_SCALE)
If DATA_TYPE <> 0 Then: YDATA_VECTOR = MATRIX_PERCENT_FUNC(YDATA_VECTOR, LOG_SCALE)
NROWS = UBound(XDATA_VECTOR, 1)

'------------------------------------------------------------------------------------------
Select Case OUTPUT
'------------------------------------------------------------------------------------------
Case 0 'INDEX MODEL
'------------------------------------------------------------------------------------------
    PORT_MARKET_TIMING_FUNC = REGRESSION_LS2_FUNC(XDATA_VECTOR, YDATA_VECTOR, True, 0, 0.95, 0)
'------------------------------------------------------------------------------------------
Case 1 'SQUARED MODEL
'------------------------------------------------------------------------------------------
    ReDim XDATA_MATRIX(1 To NROWS, 1 To 2)
    For i = 1 To NROWS
        XDATA_MATRIX(i, 1) = XDATA_VECTOR(i, 1)
        XDATA_MATRIX(i, 2) = XDATA_VECTOR(i, 1) ^ 2
    Next i
    PORT_MARKET_TIMING_FUNC = REGRESSION_LS2_FUNC(XDATA_MATRIX, YDATA_VECTOR, True, 0, 0.95, 0)
'------------------------------------------------------------------------------------------
Case 2 'HENRIKSON/MERTON MODEL
'------------------------------------------------------------------------------------------
    ReDim XDATA_MATRIX(1 To NROWS, 1 To 2)
    For i = 1 To NROWS
        XDATA_MATRIX(i, 1) = XDATA_VECTOR(i, 1)
        XDATA_MATRIX(i, 2) = MAXIMUM_FUNC(0, -XDATA_VECTOR(i, 1))
    Next i
    PORT_MARKET_TIMING_FUNC = REGRESSION_LS2_FUNC(XDATA_MATRIX, YDATA_VECTOR, True, 0, 0.95, 0)
'------------------------------------------------------------------------------------------
Case Else 'UP/DOWN MODEL
'------------------------------------------------------------------------------------------
    ReDim XDATA_MATRIX(1 To NROWS, 1 To 2)
    For i = 1 To NROWS
        XDATA_MATRIX(i, 1) = MAXIMUM_FUNC(0, -XDATA_VECTOR(i, 1))
        XDATA_MATRIX(i, 2) = MAXIMUM_FUNC(0, XDATA_VECTOR(i, 1))
    Next i
    PORT_MARKET_TIMING_FUNC = REGRESSION_LS2_FUNC(XDATA_MATRIX, YDATA_VECTOR, True, 0, 0.95, 0)
'------------------------------------------------------------------------------------------
End Select
'------------------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
PORT_MARKET_TIMING_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_BULL_BEAR_FUNC
'DESCRIPTION   : Bull - Bears Model
'LIBRARY       : FINAN_PORT
'GROUP         : TIMING
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 03/16/2009
'************************************************************************************
'************************************************************************************

Function PORT_BULL_BEAR_FUNC(ByRef XDATA_RNG As Variant, _
ByRef YDATA_RNG As Variant, _
Optional ByVal THRESD_RETURN As Double = 0, _
Optional ByVal COUNT_BASIS As Double = 12, _
Optional ByVal DATA_TYPE As Integer = 0)

Dim i As Long
Dim NROWS As Long

Dim XVAL As Double
Dim YVAL As Double

Dim XNBEARS As Long
Dim XNBULLS As Long

Dim YNBEARS As Long
Dim YNBULLS As Long

Dim XDATA_VECTOR As Variant
Dim YDATA_VECTOR As Variant

Dim XBULLS_VECTOR As Variant
Dim XBEARS_VECTOR As Variant

Dim YBULLS_VECTOR As Variant
Dim YBEARS_VECTOR As Variant

Dim OLS_MATRIX As Variant
Dim TEMP_VECTOR As Variant
Dim TEMP_MATRIX As Variant

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

If DATA_TYPE <> 0 Then: XDATA_VECTOR = MATRIX_PERCENT_FUNC(XDATA_VECTOR, 0)
If DATA_TYPE <> 0 Then: YDATA_VECTOR = MATRIX_PERCENT_FUNC(YDATA_VECTOR, 0)

NROWS = UBound(XDATA_VECTOR, 1)

ReDim TEMP_VECTOR(1 To NROWS, 1 To 1)
For i = 1 To NROWS: TEMP_VECTOR(i, 1) = YDATA_VECTOR(i, 1) - XDATA_VECTOR(i, 1): Next i

XBULLS_VECTOR = PORT_BULL_FUNC(XDATA_VECTOR, XDATA_VECTOR, THRESD_RETURN)
XBEARS_VECTOR = PORT_BEAR_FUNC(XDATA_VECTOR, XDATA_VECTOR, THRESD_RETURN)

YBULLS_VECTOR = PORT_BULL_FUNC(YDATA_VECTOR, XDATA_VECTOR, THRESD_RETURN)
YBEARS_VECTOR = PORT_BEAR_FUNC(YDATA_VECTOR, XDATA_VECTOR, THRESD_RETURN)

XNBULLS = UBound(XBULLS_VECTOR, 1): XNBEARS = UBound(XBEARS_VECTOR, 1)
YNBULLS = UBound(YBULLS_VECTOR, 1): YNBEARS = UBound(YBEARS_VECTOR, 1)

ReDim TEMP_MATRIX(1 To 26, 1 To 4)

'-------------------------------------------------------------------------------

TEMP_MATRIX(1, 1) = "BASIC RISK & RETURN CALCS"
TEMP_MATRIX(1, 2) = "BENCH(Y)"
TEMP_MATRIX(1, 3) = "PORT(X)"
TEMP_MATRIX(1, 4) = "PERFORMANCE"

XVAL = 1: YVAL = 1
For i = 1 To NROWS
    XVAL = XVAL * (1 + XDATA_VECTOR(i, 1))
    YVAL = YVAL * (1 + YDATA_VECTOR(i, 1))
Next i

TEMP_MATRIX(2, 1) = "CHAIN-LINKED RETURN"
TEMP_MATRIX(2, 2) = YVAL - 1
TEMP_MATRIX(2, 3) = XVAL - 1
TEMP_MATRIX(2, 4) = TEMP_MATRIX(2, 2) - TEMP_MATRIX(2, 3)

'-------------------------------------------------------------------------------

XVAL = 1
For i = 1 To XNBULLS: XVAL = XVAL * (1 + XBULLS_VECTOR(i, 1)): Next i

YVAL = 1
For i = 1 To YNBULLS: YVAL = YVAL * (1 + YBULLS_VECTOR(i, 1)): Next i

TEMP_MATRIX(3, 1) = "BULL"
TEMP_MATRIX(3, 2) = YVAL - 1
TEMP_MATRIX(3, 3) = XVAL - 1
TEMP_MATRIX(3, 4) = TEMP_MATRIX(3, 2) - TEMP_MATRIX(3, 3)

'-------------------------------------------------------------------------------

XVAL = 1
For i = 1 To XNBEARS: XVAL = XVAL * (1 + XBEARS_VECTOR(i, 1)): Next i

YVAL = 1
For i = 1 To YNBEARS: YVAL = YVAL * (1 + YBEARS_VECTOR(i, 1)): Next i

TEMP_MATRIX(4, 1) = "BEAR"
TEMP_MATRIX(4, 2) = YVAL - 1
TEMP_MATRIX(4, 3) = XVAL - 1
TEMP_MATRIX(4, 4) = TEMP_MATRIX(4, 2) - TEMP_MATRIX(4, 3)


'-------------------------------------------------------------------------------

TEMP_MATRIX(5, 1) = "GEOMETRIC RETURN"
TEMP_MATRIX(5, 2) = (TEMP_MATRIX(2, 2) + 1) ^ (1 / NROWS) - 1
TEMP_MATRIX(5, 3) = (TEMP_MATRIX(2, 3) + 1) ^ (1 / NROWS) - 1
TEMP_MATRIX(5, 4) = TEMP_MATRIX(5, 2) - TEMP_MATRIX(5, 3)

TEMP_MATRIX(6, 1) = "BULL"
TEMP_MATRIX(6, 2) = (TEMP_MATRIX(3, 2) + 1) ^ (1 / YNBULLS) - 1
TEMP_MATRIX(6, 3) = (TEMP_MATRIX(3, 3) + 1) ^ (1 / XNBULLS) - 1
TEMP_MATRIX(6, 4) = TEMP_MATRIX(6, 2) - TEMP_MATRIX(6, 3)

TEMP_MATRIX(7, 1) = "BEAR"
TEMP_MATRIX(7, 2) = (TEMP_MATRIX(4, 2) + 1) ^ (1 / YNBEARS) - 1
TEMP_MATRIX(7, 3) = (TEMP_MATRIX(4, 3) + 1) ^ (1 / XNBEARS) - 1
TEMP_MATRIX(7, 4) = TEMP_MATRIX(7, 2) - TEMP_MATRIX(7, 3)

'-------------------------------------------------------------------------------

TEMP_MATRIX(8, 1) = "NOBS"
TEMP_MATRIX(8, 2) = NROWS
TEMP_MATRIX(8, 3) = NROWS
TEMP_MATRIX(8, 4) = ""

TEMP_MATRIX(9, 1) = "BULL"
TEMP_MATRIX(9, 2) = YNBULLS
TEMP_MATRIX(9, 3) = XNBULLS
TEMP_MATRIX(9, 4) = ""

TEMP_MATRIX(10, 1) = "BEAR"
TEMP_MATRIX(10, 2) = YNBEARS
TEMP_MATRIX(10, 3) = XNBEARS
TEMP_MATRIX(10, 4) = ""

'-------------------------------------------------------------------------------

TEMP_MATRIX(11, 1) = "ANN. RETURN"
TEMP_MATRIX(11, 2) = ((1 + TEMP_MATRIX(2, 2)) ^ (COUNT_BASIS / TEMP_MATRIX(8, 2)) - 1)
TEMP_MATRIX(11, 3) = ((1 + TEMP_MATRIX(2, 3)) ^ (COUNT_BASIS / TEMP_MATRIX(8, 3)) - 1)
TEMP_MATRIX(11, 4) = TEMP_MATRIX(11, 2) - TEMP_MATRIX(11, 3)

TEMP_MATRIX(12, 1) = "BULL"
TEMP_MATRIX(12, 2) = ((1 + TEMP_MATRIX(3, 2)) ^ (COUNT_BASIS / TEMP_MATRIX(9, 2)) - 1)
TEMP_MATRIX(12, 3) = ((1 + TEMP_MATRIX(3, 3)) ^ (COUNT_BASIS / TEMP_MATRIX(9, 3)) - 1)
TEMP_MATRIX(12, 4) = TEMP_MATRIX(12, 2) - TEMP_MATRIX(12, 3)

TEMP_MATRIX(13, 1) = "BEAR"
TEMP_MATRIX(13, 2) = ((1 + TEMP_MATRIX(4, 2)) ^ (COUNT_BASIS / TEMP_MATRIX(10, 2)) - 1)
TEMP_MATRIX(13, 3) = ((1 + TEMP_MATRIX(4, 3)) ^ (COUNT_BASIS / TEMP_MATRIX(10, 3)) - 1)
TEMP_MATRIX(13, 4) = TEMP_MATRIX(13, 2) - TEMP_MATRIX(13, 3)


'-------------------------------------------------------------------------------

TEMP_MATRIX(14, 1) = "STDEV"
TEMP_MATRIX(14, 2) = MATRIX_STDEVP_FUNC(YDATA_VECTOR)(1, 1)
TEMP_MATRIX(14, 3) = MATRIX_STDEVP_FUNC(XDATA_VECTOR)(1, 1)
TEMP_MATRIX(14, 4) = MATRIX_STDEVP_FUNC(TEMP_VECTOR)(1, 1)

TEMP_MATRIX(15, 1) = "BULL"
TEMP_MATRIX(15, 2) = MATRIX_STDEVP_FUNC(YBULLS_VECTOR)(1, 1)
TEMP_MATRIX(15, 3) = MATRIX_STDEVP_FUNC(XBULLS_VECTOR)(1, 1)
TEMP_MATRIX(15, 4) = MATRIX_STDEVP_FUNC(PORT_BULL_FUNC(TEMP_VECTOR, XDATA_VECTOR))(1, 1)

TEMP_MATRIX(16, 1) = "BEAR"
TEMP_MATRIX(16, 2) = MATRIX_STDEVP_FUNC(YBEARS_VECTOR)(1, 1)
TEMP_MATRIX(16, 3) = MATRIX_STDEVP_FUNC(XBEARS_VECTOR)(1, 1)
TEMP_MATRIX(16, 4) = MATRIX_STDEVP_FUNC(PORT_BEAR_FUNC(TEMP_VECTOR, XDATA_VECTOR))(1, 1)
'-------------------------------------------------------------------------------
TEMP_MATRIX(17, 1) = "VOLATILITY"
TEMP_MATRIX(17, 2) = TEMP_MATRIX(14, 2) * Sqr(COUNT_BASIS)
TEMP_MATRIX(17, 3) = TEMP_MATRIX(14, 3) * Sqr(COUNT_BASIS)
TEMP_MATRIX(17, 4) = TEMP_MATRIX(14, 4) * Sqr(COUNT_BASIS)

TEMP_MATRIX(18, 1) = "BULL"
TEMP_MATRIX(18, 2) = TEMP_MATRIX(15, 2) * Sqr(COUNT_BASIS)
TEMP_MATRIX(18, 3) = TEMP_MATRIX(15, 3) * Sqr(COUNT_BASIS)
TEMP_MATRIX(18, 4) = TEMP_MATRIX(15, 4) * Sqr(COUNT_BASIS)
'TE (performance volatility)

TEMP_MATRIX(19, 1) = "BEAR"
TEMP_MATRIX(19, 2) = TEMP_MATRIX(16, 2) * Sqr(COUNT_BASIS)
TEMP_MATRIX(19, 3) = TEMP_MATRIX(16, 3) * Sqr(COUNT_BASIS)
TEMP_MATRIX(19, 4) = TEMP_MATRIX(16, 4) * Sqr(COUNT_BASIS)
'TE (performance volatility)
        
'-------------------------------------------------------------------------------
        
TEMP_MATRIX(20, 1) = ""
TEMP_MATRIX(20, 2) = ""
TEMP_MATRIX(20, 3) = "BULL"
TEMP_MATRIX(20, 4) = "BEAR"
        
TEMP_MATRIX(21, 1) = "BETA"
TEMP_MATRIX(22, 1) = "ALPHA-PERIOD"

OLS_MATRIX = REGRESSION_SIMPLE_COEF_FUNC(XDATA_VECTOR, YDATA_VECTOR)
TEMP_MATRIX(21, 2) = OLS_MATRIX(1, 1)
TEMP_MATRIX(22, 2) = OLS_MATRIX(2, 1)

OLS_MATRIX = REGRESSION_SIMPLE_COEF_FUNC(XBULLS_VECTOR, YBULLS_VECTOR)
TEMP_MATRIX(21, 3) = OLS_MATRIX(1, 1)
TEMP_MATRIX(22, 3) = OLS_MATRIX(2, 1)

OLS_MATRIX = REGRESSION_SIMPLE_COEF_FUNC(XBEARS_VECTOR, YBEARS_VECTOR)
TEMP_MATRIX(21, 4) = OLS_MATRIX(1, 1)
TEMP_MATRIX(22, 4) = OLS_MATRIX(2, 1)
                                                
'-------------------------------------------------------------------------------
                        
TEMP_MATRIX(23, 1) = "ALPHA-ANNUAL"
TEMP_MATRIX(23, 2) = (1 + TEMP_MATRIX(22, 2)) ^ COUNT_BASIS - 1
TEMP_MATRIX(23, 3) = (1 + TEMP_MATRIX(22, 3)) ^ COUNT_BASIS - 1
TEMP_MATRIX(23, 4) = (1 + TEMP_MATRIX(22, 4)) ^ COUNT_BASIS - 1
        
'-------------------------------------------------------------------------------
        
TEMP_MATRIX(24, 1) = "RHO"
TEMP_MATRIX(24, 2) = CORRELATION_FUNC(XDATA_VECTOR, YDATA_VECTOR, 0, 0)
TEMP_MATRIX(24, 3) = CORRELATION_FUNC(YBULLS_VECTOR, XBULLS_VECTOR, 0, 0)
TEMP_MATRIX(24, 4) = CORRELATION_FUNC(YBEARS_VECTOR, XBEARS_VECTOR, 0, 0)
'-------------------------------------------------------------------------------
TEMP_MATRIX(25, 1) = "R-SQR"
TEMP_MATRIX(25, 2) = TEMP_MATRIX(24, 2) ^ 2
TEMP_MATRIX(25, 3) = TEMP_MATRIX(24, 3) ^ 2
TEMP_MATRIX(25, 4) = TEMP_MATRIX(24, 4) ^ 2
        
'-------------------------------------------------------------------------------
TEMP_MATRIX(26, 1) = "TE (RESIDUAL_RISK)"
TEMP_MATRIX(26, 2) = TEMP_MATRIX(17, 2) * Sqr(1 - TEMP_MATRIX(25, 2))
TEMP_MATRIX(26, 3) = TEMP_MATRIX(18, 2) * Sqr(1 - TEMP_MATRIX(25, 3))
TEMP_MATRIX(26, 4) = TEMP_MATRIX(19, 2) * Sqr(1 - TEMP_MATRIX(25, 4))
            
'-------------------------------------------------------------------------------
            
PORT_BULL_BEAR_FUNC = TEMP_MATRIX
'-------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
PORT_BULL_BEAR_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_BULL_FUNC
'DESCRIPTION   : Extracts bull returns
'LIBRARY       : FINAN_PORT
'GROUP         : TIMING
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Private Function PORT_BULL_FUNC(ByRef XDATA_RNG As Variant, _
ByRef YDATA_RNG As Variant, _
Optional ByVal THRESD_RETURN As Double = 0)

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim NBULLS As Long

Dim XDATA_VECTOR As Variant
Dim YDATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

XDATA_VECTOR = XDATA_RNG
If UBound(XDATA_VECTOR, 1) = 1 Then: XDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(XDATA_VECTOR)

YDATA_VECTOR = YDATA_RNG
If UBound(YDATA_VECTOR, 1) = 1 Then: YDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(YDATA_VECTOR)

If UBound(XDATA_VECTOR, 1) <> UBound(YDATA_VECTOR, 1) Then: GoTo ERROR_LABEL

NROWS = UBound(YDATA_VECTOR, 1)
NBULLS = 0
For i = 1 To NROWS
    If YDATA_VECTOR(i, 1) >= THRESD_RETURN Then: NBULLS = NBULLS + 1
Next i

ReDim TEMP_VECTOR(1 To NBULLS, 1 To 1)
j = 1
For i = 1 To NROWS
    If YDATA_VECTOR(i, 1) >= THRESD_RETURN Then
        TEMP_VECTOR(j, 1) = XDATA_VECTOR(i, 1)
        j = j + 1
    End If
Next i

PORT_BULL_FUNC = TEMP_VECTOR
    
Exit Function
ERROR_LABEL:
PORT_BULL_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_BEAR_FUNC
'DESCRIPTION   : Extracts bear returns
'LIBRARY       : FINAN_PORT
'GROUP         : TIMING
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Private Function PORT_BEAR_FUNC(ByRef XDATA_RNG As Variant, _
ByRef YDATA_RNG As Variant, _
Optional ByVal THRESD_RETURN As Double = 0)

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim NBEARS As Long

Dim XDATA_VECTOR As Variant
Dim YDATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

XDATA_VECTOR = XDATA_RNG
If UBound(XDATA_VECTOR, 1) = 1 Then: XDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(XDATA_VECTOR)

YDATA_VECTOR = YDATA_RNG
If UBound(YDATA_VECTOR, 1) = 1 Then: YDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(YDATA_VECTOR)

If UBound(XDATA_VECTOR, 1) <> UBound(YDATA_VECTOR, 1) Then: GoTo ERROR_LABEL

NROWS = UBound(YDATA_VECTOR, 1)
NBEARS = 0
For i = 1 To NROWS
    If YDATA_VECTOR(i, 1) < THRESD_RETURN Then: NBEARS = NBEARS + 1
Next i

ReDim TEMP_VECTOR(1 To NBEARS, 1 To 1)
j = 1
For i = 1 To NROWS
    If YDATA_VECTOR(i, 1) < THRESD_RETURN Then
        TEMP_VECTOR(j, 1) = XDATA_VECTOR(i, 1)
        j = j + 1
    End If
Next i

PORT_BEAR_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
PORT_BEAR_FUNC = Err.number
End Function
