Attribute VB_Name = "FINAN_PORT_SIMUL_CPPI_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_CPPI_SIMULATION_FUNC

'DESCRIPTION   : Constant Proportion Portfolio Insurance (CPPI) Simulator featuring
'realistic stochastics, i.e. negative skewness (long left tails) and
'excess kurtosis (fat tails). The spreadsheet does not contain any
'explanations of the CPPI strategy, it assumes that you are familiar
'with CCPI concepts like "Multiplier", "Cushion" and so on. The realism
'in simulating the risky asset is achieved with some help from my NIG
'(Normal Inverse Gaussian). An interesting extensions would be VPPI:
'Variable proportion portfolio insurance, i.e. a dynamic protective
'strategy considerung changes in the risk characteristics over time.

'http://en.wikipedia.org/wiki/Constant_proportion_portfolio_insurance

'LIBRARY       : PORTFOLIO
'GROUP         : SIMULATION_CPPI
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************

Function PORT_CPPI_SIMULATION_FUNC( _
Optional ByVal EXPECTED_RETURN As Double = 0.15, _
Optional ByVal VOLATILITY As Double = 0.3, _
Optional ByVal SKEWNESS As Double = -1, _
Optional ByVal KURTOSIS As Double = 5, _
Optional ByVal RISKFREE_RATE As Double = 0.02, _
Optional ByVal BEGINNING_NAV As Double = 100, _
Optional ByVal MULTIPLIER As Double = 5, _
Optional ByVal PROTECTION_LEVEL As Double = 0.9, _
Optional ByVal MIN_RISKY_EXPOSURE As Double = -0.3, _
Optional ByVal MAX_RISKY_EXPOSURE As Double = 1.3, _
Optional ByVal TRANSACTION_COST_BP As Double = 10, _
Optional ByVal COUNT_BASIS As Long = 250, _
Optional ByVal nLOOPS As Long = 250, _
Optional ByVal OUTPUT As Integer = 0)

'SKEWNESS (<0: longer left tail)
'KURTOSIS (3=Normal Distr)

Dim i As Long
Dim TIME_STEP As Double
Dim ATEMP_SUM As Double
Dim BTEMP_SUM As Double

Dim TEMP_MATRIX As Variant
Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL

TIME_STEP = 1 / COUNT_BASIS

PARAM_VECTOR = NIG_MLE_PARAMETERS_FUNC(0, 1, SKEWNESS, KURTOSIS + 3)

ReDim TEMP_MATRIX(0 To nLOOPS + 1, 1 To 14)

TEMP_MATRIX(0, 1) = "TIME INDEX"
TEMP_MATRIX(0, 2) = "STOCHASTICS"
TEMP_MATRIX(0, 3) = "RISKY STRATEGY"
TEMP_MATRIX(0, 4) = "RETURN RISKY STRATEGY"
TEMP_MATRIX(0, 5) = "RISKFREE STRATEGY"
TEMP_MATRIX(0, 6) = "RETURN RISKFREE STRATEGY"
TEMP_MATRIX(0, 7) = "PROTECTION LEVEL"
TEMP_MATRIX(0, 8) = "FLOOR"
TEMP_MATRIX(0, 9) = "CPPI STRATEGY"
TEMP_MATRIX(0, 10) = "VIOLATION"
TEMP_MATRIX(0, 11) = "CUSHION"
TEMP_MATRIX(0, 12) = "RISKY EXPOSURE"
TEMP_MATRIX(0, 13) = "RISKFREE EXPOSURE"
TEMP_MATRIX(0, 14) = "TURNOVER"

TEMP_MATRIX(1, 1) = 0
TEMP_MATRIX(1, 2) = 0
TEMP_MATRIX(1, 3) = BEGINNING_NAV
TEMP_MATRIX(1, 4) = ""
TEMP_MATRIX(1, 5) = BEGINNING_NAV
TEMP_MATRIX(1, 6) = ""
TEMP_MATRIX(1, 7) = BEGINNING_NAV * PROTECTION_LEVEL

TEMP_MATRIX(1, 8) = PROTECTION_LEVEL * BEGINNING_NAV * _
                    Exp(-RISKFREE_RATE * (nLOOPS - _
                    TEMP_MATRIX(1, 1)) * TIME_STEP)

TEMP_MATRIX(1, 9) = BEGINNING_NAV

TEMP_MATRIX(1, 10) = IIf(TEMP_MATRIX(1, 9) < TEMP_MATRIX(1, 7), 1, 0)
TEMP_MATRIX(1, 11) = MAXIMUM_FUNC(TEMP_MATRIX(1, 9) - TEMP_MATRIX(1, 8), 0)

TEMP_MATRIX(1, 12) = _
MAXIMUM_FUNC(MINIMUM_FUNC(MULTIPLIER * TEMP_MATRIX(1, 11) / 100, _
MAX_RISKY_EXPOSURE), MIN_RISKY_EXPOSURE)

TEMP_MATRIX(1, 13) = 1 - TEMP_MATRIX(1, 12)
TEMP_MATRIX(1, 14) = 1

ATEMP_SUM = TEMP_MATRIX(1, 10)
BTEMP_SUM = TEMP_MATRIX(1, 14)

'------------------------------------------------------------------------------
For i = 1 To nLOOPS
'------------------------------------------------------------------------------
    TEMP_MATRIX(i + 1, 1) = i
    TEMP_MATRIX(i + 1, 2) = NIG_RANDOM_FUNC(PARAM_VECTOR(1, 1), _
                            PARAM_VECTOR(2, 1), PARAM_VECTOR(3, 1), _
                            PARAM_VECTOR(4, 1))
    
    TEMP_MATRIX(i + 1, 3) = TEMP_MATRIX(i, 3) * (Exp(EXPECTED_RETURN * TIME_STEP + _
                            VOLATILITY * TEMP_MATRIX(i + 1, 2) * Sqr(TIME_STEP)))
    
    TEMP_MATRIX(i + 1, 4) = TEMP_MATRIX(i + 1, 3) / TEMP_MATRIX(i, 3) - 1
    TEMP_MATRIX(i + 1, 5) = TEMP_MATRIX(i, 5) * Exp(RISKFREE_RATE * TIME_STEP)
    TEMP_MATRIX(i + 1, 6) = TEMP_MATRIX(i + 1, 5) / TEMP_MATRIX(i, 5) - 1
    
    TEMP_MATRIX(i + 1, 7) = BEGINNING_NAV * PROTECTION_LEVEL
    
    TEMP_MATRIX(i + 1, 8) = PROTECTION_LEVEL * BEGINNING_NAV * _
                    Exp(-RISKFREE_RATE * (nLOOPS - _
                    TEMP_MATRIX(i + 1, 1)) * TIME_STEP)
    
    
    '---------------------------------------------------------------------------
    TEMP_MATRIX(i + 1, 9) = TEMP_MATRIX(i, 9) * (1 + TEMP_MATRIX(i, 12) * _
                            TEMP_MATRIX(i + 1, 4) + TEMP_MATRIX(i, 13) * _
                            TEMP_MATRIX(i + 1, 6)) - TEMP_MATRIX(i, 14) * _
                            TRANSACTION_COST_BP / 100
    
    If TEMP_MATRIX(i + 1, 9) < TEMP_MATRIX(i + 1, 7) Then
        TEMP_MATRIX(i + 1, 10) = 1
    Else
        TEMP_MATRIX(i + 1, 10) = 0
    End If
    
    TEMP_MATRIX(i + 1, 11) = _
    MAXIMUM_FUNC(TEMP_MATRIX(i + 1, 9) - TEMP_MATRIX(i + 1, 8), 0)
    
    TEMP_MATRIX(i + 1, 12) = _
    MAXIMUM_FUNC(MINIMUM_FUNC(MULTIPLIER * TEMP_MATRIX(i + 1, 11) / 100, _
    MAX_RISKY_EXPOSURE), MIN_RISKY_EXPOSURE)
    
    
    TEMP_MATRIX(i + 1, 13) = 1 - TEMP_MATRIX(i + 1, 12)
    
    TEMP_MATRIX(i + 1, 14) = Abs(TEMP_MATRIX(i + 1, 12) - TEMP_MATRIX(i, 12))
    
    ATEMP_SUM = ATEMP_SUM + TEMP_MATRIX(i + 1, 10)

    BTEMP_SUM = BTEMP_SUM + TEMP_MATRIX(i + 1, 14)
'------------------------------------------------------------------------------
Next i
'------------------------------------------------------------------------------

Select Case OUTPUT
Case 0
    PORT_CPPI_SIMULATION_FUNC = TEMP_MATRIX
Case Else
    'Protection Level Violations, Terminal Value is Violation, Total Turnover
    PORT_CPPI_SIMULATION_FUNC = Array(ATEMP_SUM, TEMP_MATRIX(nLOOPS + 1, 9) < _
                                TEMP_MATRIX(nLOOPS + 1, 7), BTEMP_SUM)
End Select

Exit Function
ERROR_LABEL:
PORT_CPPI_SIMULATION_FUNC = Err.number
End Function
