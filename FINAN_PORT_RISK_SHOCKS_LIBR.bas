Attribute VB_Name = "FINAN_PORT_RISK_SHOCKS_LIBR"

'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------

'************************************************************************************
'************************************************************************************
'FUNCTION      : RNG_PORT_SHOCKS_SCENARIOS_FUNC

'DESCRIPTION   : MTM Model of portfolio; The Factor-Based approach to calculating
'VAR begins with a principal components analysis of the yield curve. This decomposes
'yield curve movements into a small number of underlying factors including a “Shift”
'factor that allows rates to rise or fall and a “Twist” factor that allows the curve
'to steepen or flatten. Combining these factors produces specific yield curve
'scenarios used to estimate hypothetical portfolio profit or loss. The greatest
'loss among these scenarios provides an intuitive and rapid VAR estimate that tends
'to provide a conservative estimate of the nominal percentile of the loss distribution.

'LIBRARY       : PORT_RISK
'GROUP         : SHOCKS
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'UPDATE        : 07/27/2008
'************************************************************************************
'************************************************************************************

Function RNG_PORT_SHOCKS_SCENARIOS_FUNC(ByRef DST_RNG As Excel.Range, _
ByVal NO_FACTORS As Long, _
ByVal NO_ASSETS As Long, _
Optional ByVal DENOM_VAL As Double = 100, _
Optional ByVal RNG_NAME_FLAG As Boolean = False)

Dim i As Long
Dim j As Long
Dim k As Long

Dim SHOCKS_RNG As Excel.Range
Dim FACTORS_RNG As Excel.Range
Dim MARKET_RNG As Excel.Range
Dim SCENARIO_RNG As Excel.Range

Dim QUANTITY_RNG As Excel.Range
Dim STRIKE_RNG As Excel.Range
Dim SIGMA_RNG As Excel.Range
Dim EXPIRATION_RNG As Excel.Range
Dim FLAG_RNG As Excel.Range

Dim PORT_RNG As Excel.Range

On Error GoTo ERROR_LABEL

RNG_PORT_SHOCKS_SCENARIOS_FUNC = False

k = 5
'--------------------------------FIRST PASS: SHOCKS MATRIX----------------------------

DST_RNG.Offset(-1, 0).value = "SHOCKS MATRIX"
DST_RNG.Offset(-1, 0).Font.Bold = True

Set SHOCKS_RNG = Range(DST_RNG.Offset(1, 1), DST_RNG.Offset(NO_FACTORS, NO_ASSETS))
If RNG_NAME_FLAG = True Then: SHOCKS_RNG.name = "SHOCKS_MATRIX"

SHOCKS_RNG.value = 0
SHOCKS_RNG.Font.ColorIndex = 5

For i = 1 To NO_FACTORS
    SHOCKS_RNG.Cells(i, 0) = "SHOCK " & CStr(i)
    SHOCKS_RNG.Cells(i, 0).Font.ColorIndex = 3
Next i

For j = 1 To NO_ASSETS
    SHOCKS_RNG.Cells(0, j) = "PCA - " & CStr(j)
    SHOCKS_RNG.Cells(0, j).Font.ColorIndex = 3
Next j

Set DST_RNG = DST_RNG.Offset(NO_FACTORS + k, 0)


'--------------------------------SECOND PASS: FACTORS & MARKET ----------------------------

DST_RNG.Offset(-2, 0).value = "FACTORS VECTORS"
DST_RNG.Offset(-2, 0).Font.Bold = True

DST_RNG.value = "FACTORS"
DST_RNG.Offset(1, 0).value = "MARKET"

Set FACTORS_RNG = Range(DST_RNG.Offset(0, 1), DST_RNG.Offset(0, NO_ASSETS))
If RNG_NAME_FLAG = True Then: FACTORS_RNG.name = "FACTORS_RNG"

FACTORS_RNG.value = 0
FACTORS_RNG.Font.ColorIndex = 5

For i = 1 To NO_ASSETS
    FACTORS_RNG.Cells(0, i).formula = "=" & SHOCKS_RNG.Cells(0, i).Address
Next i

Set MARKET_RNG = Range(DST_RNG.Offset(1, 1), DST_RNG.Offset(1, NO_ASSETS))
If RNG_NAME_FLAG = True Then: MARKET_RNG.name = "MARKET_RNG"

MARKET_RNG.value = 0
MARKET_RNG.Font.ColorIndex = 5

Set DST_RNG = DST_RNG.Offset(k, 0)

'--------------------------------THIRD PASS: MARKET_SCENARIO ----------------------------

Range(DST_RNG, DST_RNG.Offset(2, NO_FACTORS)).FormulaArray = "=PORT_SHOCKS_SCENARIOS_FUNC(" & SHOCKS_RNG.Address & "," & FACTORS_RNG.Address & "," & MARKET_RNG.Address & "," & DENOM_VAL & ")"
Set SCENARIO_RNG = Range(DST_RNG.Offset(2, 1), DST_RNG.Offset(2, NO_FACTORS))
If RNG_NAME_FLAG = True Then: SCENARIO_RNG.name = "SCENARIO_RNG"

For i = 1 To NO_FACTORS
    SCENARIO_RNG.Cells(-2, i).formula = "=" & SHOCKS_RNG.Cells(i, 0).Address
Next i

Set DST_RNG = DST_RNG.Offset(k + 2, 0)

'--------------------------------FORTH PASS: OPT_PORTFOLIO----------------------------


DST_RNG.Offset(-2, 0).value = "PORTFOLIO MARKET VALUE"
DST_RNG.Offset(-2, 0).Font.Bold = True

DST_RNG.value = "QUANTITY"
DST_RNG.Offset(1, 0).value = "STRIKE"
DST_RNG.Offset(2, 0).value = "SIGMA"
DST_RNG.Offset(3, 0).value = "EXPIRATION"
DST_RNG.Offset(4, 0).value = "OPTION FLAG"
DST_RNG.Offset(6, 0).value = "PORTFOLIO MARKET VALUE"

DST_RNG.Offset(6, 0).Font.Bold = True

Set QUANTITY_RNG = Range(DST_RNG.Offset(0, 1), DST_RNG.Offset(0, NO_FACTORS))
For i = 1 To NO_FACTORS
    QUANTITY_RNG.Cells(0, i).formula = "=" & SHOCKS_RNG.Cells(i, 0).Address
Next i

Set STRIKE_RNG = Range(DST_RNG.Offset(1, 1), DST_RNG.Offset(1, NO_FACTORS))
Set SIGMA_RNG = Range(DST_RNG.Offset(2, 1), DST_RNG.Offset(2, NO_FACTORS))
Set EXPIRATION_RNG = Range(DST_RNG.Offset(3, 1), DST_RNG.Offset(3, NO_FACTORS))
Set FLAG_RNG = Range(DST_RNG.Offset(4, 1), DST_RNG.Offset(4, NO_FACTORS))
Set PORT_RNG = Range(DST_RNG.Offset(6, 1), DST_RNG.Offset(6, NO_FACTORS))

If RNG_NAME_FLAG = True Then: PORT_RNG.name = "MARKET_PORT_RNG"

QUANTITY_RNG.value = 0
QUANTITY_RNG.Font.ColorIndex = 5

STRIKE_RNG.value = 0
STRIKE_RNG.Font.ColorIndex = 5

SIGMA_RNG.value = 0
SIGMA_RNG.Font.ColorIndex = 5

EXPIRATION_RNG.value = 0
EXPIRATION_RNG.Font.ColorIndex = 5

FLAG_RNG.value = 1
FLAG_RNG.Font.ColorIndex = 3

With FLAG_RNG.Validation
    .Delete
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
    xlBetween, Formula1:="1,-1"
    .IgnoreBlank = True
    .InCellDropdown = True
    .InputTitle = ""
    .ErrorTitle = ""
    .InputMessage = ""
    .ErrorMessage = ""
    .ShowInput = True
    .ShowError = True
End With

PORT_RNG.Font.Bold = True

PORT_RNG.FormulaArray = "=PORT_SHOCKS_MARKET_VALUE_FUNC(" & QUANTITY_RNG.Address & "," & _
    SCENARIO_RNG.Address & "," & _
    STRIKE_RNG.Address & "," & _
    SIGMA_RNG.Address & "," & _
    EXPIRATION_RNG.Address & "," & _
    FLAG_RNG.Address & ")"

RNG_PORT_SHOCKS_SCENARIOS_FUNC = True

Exit Function
ERROR_LABEL:
RNG_PORT_SHOCKS_SCENARIOS_FUNC = False
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_SHOCKS_SCENARIOS_FUNC
'DESCRIPTION   : This function accept random drivers (appropriate for Monte
'Carlo simulation), or deterministic drivers (appropriate for filling
'in the "grid values" for scenario simulation).
'LIBRARY       : PORT_RISK
'GROUP         : SHOCKS
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'UPDATE        : 07/27/2008
'************************************************************************************
'************************************************************************************

Function PORT_SHOCKS_SCENARIOS_FUNC(ByRef SHOCKS_RNG As Variant, _
ByRef FACTORS_RNG As Variant, _
ByRef MARKET_RNG As Variant, _
Optional ByVal SCALAR As Double = 100)

'SHOCKS_RNG = Factor Table (Case 0) -->  Function

'---------------------------------------------------------------------------------
'---------------------------------------------------------------------------------

'Any continuous function can be approximated arbitrarily well by a piecewise
'linear function. As the number of regularly spaced intervals increases,
'the piecewise linear function grows closer to the continuous function it
'approximates. As a practical matter, one needs only a few intervals to
'approximate the value of non-exotic derivatives. (Piecewise approximation may
'not be appropriate if the portfolio has considerable exposure to deals such
'as digital options.)

'----------------------------------------------------------------------------------
'----------------------------------------------------------------------------------

Dim i As Long
Dim NCOLUMNS As Long

Dim DATA_MATRIX As Variant
Dim SHOCKS_MATRIX As Variant
Dim FACTORS_VECTOR As Variant
Dim MARKET_VECTOR As Variant

On Error GoTo ERROR_LABEL

SHOCKS_MATRIX = SHOCKS_RNG

FACTORS_VECTOR = FACTORS_RNG
If UBound(FACTORS_VECTOR, 2) = 1 Then
    FACTORS_VECTOR = MATRIX_TRANSPOSE_FUNC(FACTORS_VECTOR)
End If
If UBound(SHOCKS_MATRIX, 2) <> UBound(FACTORS_VECTOR, 2) Then: GoTo ERROR_LABEL

MARKET_VECTOR = MARKET_RNG
If UBound(MARKET_VECTOR, 2) = 1 Then
    MARKET_VECTOR = MATRIX_TRANSPOSE_FUNC(MARKET_VECTOR)
End If
DATA_MATRIX = MMULT_FUNC(FACTORS_VECTOR, MATRIX_TRANSPOSE_FUNC(SHOCKS_MATRIX))
NCOLUMNS = UBound(DATA_MATRIX, 2)

ReDim TEMP_MATRIX(1 To 3, 1 To NCOLUMNS + 1)

For i = 1 To NCOLUMNS
    TEMP_MATRIX(1, i + 1) = MARKET_VECTOR(1, i)
    TEMP_MATRIX(2, i + 1) = DATA_MATRIX(1, i)
    TEMP_MATRIX(3, i + 1) = TEMP_MATRIX(1, i + 1) + TEMP_MATRIX(2, i + 1) / SCALAR
Next i

TEMP_MATRIX(1, 1) = ("BASE MARKET LEVEL")
TEMP_MATRIX(2, 1) = ("MARKET SCENARIO (IN BASIS POINTS)")
TEMP_MATRIX(3, 1) = ("MARKET SCENARIO")

PORT_SHOCKS_SCENARIOS_FUNC = TEMP_MATRIX 'Market scenario --> MARKET_RNG in the
'Market Value of the Portfolio

Exit Function
ERROR_LABEL:
PORT_SHOCKS_SCENARIOS_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_SHOCKS_SIMULATION_FUNC
'DESCRIPTION   : Portfolio Shocks Scenario Analysis
'LIBRARY       : PORT_RISK
'GROUP         : SHOCKS
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'UPDATE        : 07/27/2008
'************************************************************************************
'************************************************************************************

Function PORT_SHOCKS_SIMULATION_FUNC(ByVal nLOOPS As Long, _
ByRef DATA_RNG As Variant, _
ByVal NO_SHOCKS As Long, _
Optional ByVal OUTPUT As Integer = 1, _
Optional ByVal DATA_TYPE As Integer = 1, _
Optional ByVal LOG_SCALE As Integer = 0, _
Optional ByVal RANDOM_FLAG As Boolean = True, _
Optional ByVal MOMENTS_FLAG As Boolean = True, _
Optional ByRef CORREL_RNG As Variant)

'The portfolio scenario starts with a scenario for the drivers, expressed
'as a number of standard deviations of a market factor.

Dim i As Long
Dim j As Long
Dim k As Long

Dim NSIZE As Long

Dim SHOCKS_MATRIX As Variant
Dim SHIFTS_MATRIX As Variant
Dim NORMAL_RANDOM_MATRIX As Variant

Dim TEMP_SUM As Variant
Dim SORTED_EIGEN_VALUES As Variant
Dim LARGEST_EIGEN_VECTORS As Variant

On Error GoTo ERROR_LABEL

'-----------------------------------------------------------------------
'----------------Extracting the Column and Deleting Headings
'-----------------------------------------------------------------------
SORTED_EIGEN_VALUES = MATRIX_GET_COLUMN_FUNC(MATRIX_REMOVE_ROWS_FUNC(MATRIX_PCA_FACTORS_FUNC(DATA_RNG, DATA_TYPE, LOG_SCALE, 2), 1, 1), 1, 1) 'Sorted Eigen Values
LARGEST_EIGEN_VECTORS = MATRIX_REMOVE_ROWS_FUNC(MATRIX_PCA_FACTORS_FUNC(DATA_RNG, DATA_TYPE, LOG_SCALE, 1), 1, 1) 'Largest Eigen Vectors
'--------------------------------------------------------------------------------------
If NO_SHOCKS > UBound(SORTED_EIGEN_VALUES, 1) Then: NO_SHOCKS = UBound(SORTED_EIGEN_VALUES, 1)
'--------------------------------------------------------------------------------------
NSIZE = UBound(LARGEST_EIGEN_VECTORS, 1)

If RANDOM_FLAG = True Then: Randomize

NORMAL_RANDOM_MATRIX = MULTI_NORMAL_RANDOM_MATRIX_FUNC(0, nLOOPS, NSIZE, 0, 1, RANDOM_FLAG, MOMENTS_FLAG, CORREL_RNG, 0)
ReDim SHIFTS_MATRIX(1 To NSIZE, 1 To nLOOPS)
For j = 1 To nLOOPS
    For i = 1 To NSIZE
        TEMP_SUM = 0
        For k = 1 To NO_SHOCKS
            TEMP_SUM = TEMP_SUM + (SORTED_EIGEN_VALUES(k, 1) ^ 0.5 * NORMAL_RANDOM_MATRIX(j, k)) * LARGEST_EIGEN_VECTORS(i, k)
        Next k
        SHIFTS_MATRIX(i, j) = TEMP_SUM
    Next i
Next j
    
SHIFTS_MATRIX = DATA_BASIC_MOMENTS_FUNC(MATRIX_TRANSPOSE_FUNC(SHIFTS_MATRIX), 0, 0, 0.05, 1)
SHIFTS_MATRIX = MATRIX_ADD_COLUMNS_FUNC(SHIFTS_MATRIX, 1, 1)
        
SHIFTS_MATRIX(1, 1) = "SIMULATED FACTOR SHOCKS %-SHIFT SUMMARY"
For j = 2 To NSIZE + 1
    SHIFTS_MATRIX(j, 1) = "Nearby: " & CStr(j - 1)
Next j
        
If OUTPUT = 0 Then
    PORT_SHOCKS_SIMULATION_FUNC = SHIFTS_MATRIX
    Exit Function
End If
        
ReDim SHOCKS_MATRIX(1 To NO_SHOCKS, 1 To nLOOPS)
For j = 1 To nLOOPS
    For k = 1 To NO_SHOCKS
        SHOCKS_MATRIX(k, j) = (SORTED_EIGEN_VALUES(k, 1) ^ 0.5 * NORMAL_RANDOM_MATRIX(j, k))
    Next k
Next j
        
SHOCKS_MATRIX = DATA_BASIC_MOMENTS_FUNC(MATRIX_TRANSPOSE_FUNC(SHOCKS_MATRIX), 0, 0, 0.05, 1)
SHOCKS_MATRIX = MATRIX_ADD_COLUMNS_FUNC(SHOCKS_MATRIX, 1, 1)
SHOCKS_MATRIX(1, 1) = "SIMULATED FACTOR SHOCKS SUMMARY"
For j = 2 To NO_SHOCKS + 1
    SHOCKS_MATRIX(j, 1) = "Shock: " & CStr(j - 1)
Next j

If OUTPUT = 1 Then
    PORT_SHOCKS_SIMULATION_FUNC = SHOCKS_MATRIX
    Exit Function
End If
    
PORT_SHOCKS_SIMULATION_FUNC = Array(SHIFTS_MATRIX, SHOCKS_MATRIX)

Exit Function
ERROR_LABEL:
PORT_SHOCKS_SIMULATION_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_SHOCKS_MARKET_VALUE_FUNC
'DESCRIPTION   : MTM Value of portfolio
'LIBRARY       : PORT_RISK
'GROUP         : SHOCKS
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'UPDATE        : 07/27/2008
'************************************************************************************
'************************************************************************************

Function PORT_SHOCKS_MARKET_VALUE_FUNC(ByRef QUANT_RNG As Variant, _
ByRef MARKET_RNG As Variant, _
ByRef STRIKE_RNG As Variant, _
ByRef SIGMA_RNG As Variant, _
ByRef EXPIRATION_RNG As Variant, _
ByRef FLAG_RNG As Variant, _
Optional ByVal OUTPUT As Integer = 0)

'QUANT_RNG --> Number of contracts
'STRIKE_RNG --> STRIKE
'FLAG_RNG --> Put(1) / Call(-1)
'SIGMA_RNG --> Volatility
'EXPIRATION_RNG --> Time
'MARKET_RNG --> Market Rate (Spot Rate)

'-------------------------------------------------------------
'KEY THING: The option pays off on the interest rate itself, not on an
'intrument which is a non-linear function of the rate.
'-------------------------------------------------------------

Dim i As Long
Dim NCOLUMNS As Long
Dim TEMP_SUM As Double

Dim TEMP_VECTOR As Variant

Dim QUANT_VECTOR As Variant
Dim MARKET_VECTOR As Variant
Dim STRIKE_VECTOR As Variant
Dim SIGMA_VECTOR As Variant
Dim EXPIRATION_VECTOR As Variant
Dim FLAG_VECTOR As Variant

On Error GoTo ERROR_LABEL

QUANT_VECTOR = QUANT_RNG
If UBound(QUANT_VECTOR, 2) = 1 Then
    QUANT_VECTOR = MATRIX_TRANSPOSE_FUNC(QUANT_VECTOR)
End If

MARKET_VECTOR = MARKET_RNG
If UBound(MARKET_VECTOR, 2) = 1 Then
    MARKET_VECTOR = MATRIX_TRANSPOSE_FUNC(MARKET_VECTOR)
End If

STRIKE_VECTOR = STRIKE_RNG
If UBound(STRIKE_VECTOR, 2) = 1 Then
    STRIKE_VECTOR = MATRIX_TRANSPOSE_FUNC(STRIKE_VECTOR)
End If

SIGMA_VECTOR = SIGMA_RNG
If UBound(SIGMA_VECTOR, 2) = 1 Then
    SIGMA_VECTOR = MATRIX_TRANSPOSE_FUNC(SIGMA_VECTOR)
End If

EXPIRATION_VECTOR = EXPIRATION_RNG
If UBound(EXPIRATION_VECTOR, 2) = 1 Then
    EXPIRATION_VECTOR = MATRIX_TRANSPOSE_FUNC(EXPIRATION_VECTOR)
End If

FLAG_VECTOR = FLAG_RNG
If UBound(FLAG_VECTOR, 2) = 1 Then
    FLAG_VECTOR = MATRIX_TRANSPOSE_FUNC(FLAG_VECTOR)
End If

NCOLUMNS = UBound(QUANT_VECTOR, 2)
ReDim TEMP_VECTOR(1 To 1, 1 To NCOLUMNS)

TEMP_SUM = 0
For i = 1 To NCOLUMNS
    Select Case FLAG_VECTOR(1, i)
    Case 1 ', "CALL", "C"
        TEMP_VECTOR(1, i) = SHOCKS_CALL_OPTION_FUNC(MARKET_VECTOR(1, i), STRIKE_VECTOR(1, i), SIGMA_VECTOR(1, i), EXPIRATION_VECTOR(1, i), 0) * QUANT_VECTOR(1, i)
    Case -1 ', "PUT", "P"
        TEMP_VECTOR(1, i) = SHOCKS_PUT_OPTION_FUNC(MARKET_VECTOR(1, i), STRIKE_VECTOR(1, i), SIGMA_VECTOR(1, i), EXPIRATION_VECTOR(1, i), 0) * QUANT_VECTOR(1, i)
    End Select
   TEMP_SUM = TEMP_SUM + TEMP_VECTOR(1, i)
Next i

Select Case OUTPUT
Case 0
    PORT_SHOCKS_MARKET_VALUE_FUNC = TEMP_VECTOR
Case Else
    PORT_SHOCKS_MARKET_VALUE_FUNC = TEMP_SUM
End Select
    
Exit Function
ERROR_LABEL:
PORT_SHOCKS_MARKET_VALUE_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : SHOCKS_PUT_OPTION_FUNC
'DESCRIPTION   : Put Value
'LIBRARY       : PORT_RISK
'GROUP         : SHOCKS
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'UPDATE        : 07/27/2008
'************************************************************************************
'************************************************************************************

Private Function SHOCKS_PUT_OPTION_FUNC(ByVal SPOT As Double, _
ByVal STRIKE As Double, _
ByVal SIGMA As Double, _
ByVal EXPIRATION As Double, _
Optional ByVal CND_TYPE As Integer = 0)

Dim D1_VAL As Double
Dim D2_VAL As Double

On Error GoTo ERROR_LABEL

D1_VAL = (Log(SPOT / STRIKE) + EXPIRATION * SIGMA ^ 2 / 2) / (SIGMA * Sqr(EXPIRATION))
D2_VAL = (Log(SPOT / STRIKE) - EXPIRATION * SIGMA ^ 2 / 2) / (SIGMA * Sqr(EXPIRATION))

SHOCKS_PUT_OPTION_FUNC = STRIKE * CND_FUNC(-D2_VAL, CND_TYPE) - SPOT * CND_FUNC(-D1_VAL, CND_TYPE)

Exit Function
ERROR_LABEL:
SHOCKS_PUT_OPTION_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : SHOCKS_CALL_OPTION_FUNC
'DESCRIPTION   : Call Value
'LIBRARY       : PORT_RISK
'GROUP         : SHOCKS
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'UPDATE        : 07/27/2008
'************************************************************************************
'************************************************************************************

Private Function SHOCKS_CALL_OPTION_FUNC(ByVal SPOT As Double, _
ByVal STRIKE As Double, _
ByVal SIGMA As Double, _
ByVal EXPIRATION As Double, _
Optional ByVal CND_TYPE As Integer = 0)

Dim D1_VAL As Double
Dim D2_VAL As Double

On Error GoTo ERROR_LABEL

D1_VAL = (Log(SPOT / STRIKE) + EXPIRATION * SIGMA ^ 2 / 2) / (SIGMA * Sqr(EXPIRATION))
D2_VAL = (Log(SPOT / STRIKE) - EXPIRATION * SIGMA ^ 2 / 2) / (SIGMA * Sqr(EXPIRATION))

SHOCKS_CALL_OPTION_FUNC = SPOT * CND_FUNC(D1_VAL, CND_TYPE) - STRIKE * CND_FUNC(D2_VAL, CND_TYPE)

Exit Function
ERROR_LABEL:
SHOCKS_CALL_OPTION_FUNC = Err.number
End Function
