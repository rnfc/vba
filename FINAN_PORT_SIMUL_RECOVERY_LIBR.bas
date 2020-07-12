Attribute VB_Name = "FINAN_PORT_SIMUL_RECOVERY_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_EXPECTED_RECOVERY_PERIOD_FUNC
'DESCRIPTION   : Calculation of expected recovery period given a time series of
'returns, a current low portfolio value and a higher portfolio value at recovery.
'LIBRARY       : PORTFOLIO
'GROUP         : SIMULATION_RECOVERY
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function PORT_EXPECTED_RECOVERY_PERIOD_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal CURRENT_LEVEL_VAL As Double = 75, _
Optional ByVal RECOVERY_LEVEL_VAL As Double = 100, _
Optional ByVal NSIZE As Long = 10000, _
Optional ByVal nLOOPS As Long = 5000)

'Current Portfolio Value --> 75
'Portfolio Value @ Recovery --> 100

'Maximum number of steps --> 10000
'Number of Simulations --> 5000

'With bootstrapping from observed returns, the expected recovery time given a current
'portfolio value and a higher target portfolio value are calculated.

'Most other recovery period calculators are based on implicit or explicit parametric
'assumptions (e.g. log-normal distribution of return). Bootstrapping on the other hand
'takes into account "black swan risks", fat tails and skewness of real-world financial data

Dim i As Long
Dim j As Long
Dim k As Long

Dim hh As Long 'RecoveryPeriod
Dim ii As Long 'nRecovery
Dim jj As Long 'avg recovery period
Dim kk As Long 'min recovery period
Dim ll As Long 'max recovery period

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_VAL As Double

Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

ReDim TEMP_MATRIX(1 To 3, 1 To NCOLUMNS + 1)
TEMP_MATRIX(1, 1) = "Expected Recovery Time"
TEMP_MATRIX(2, 1) = "Minimum Recovery Time"
TEMP_MATRIX(3, 1) = "Maximum Recovery Time"

For k = 1 To NCOLUMNS
    ii = 0: jj = 0
    kk = NSIZE
    ll = -NSIZE
    For j = 1 To nLOOPS
        TEMP_VAL = 1
        For i = 1 To NSIZE
            TEMP_VAL = TEMP_VAL * (1 + DATA_MATRIX(Int(Rnd() * NROWS) + 1, 1))
            If CURRENT_LEVEL_VAL * TEMP_VAL >= RECOVERY_LEVEL_VAL Then
                ii = ii + 1
                hh = hh + i
                kk = IIf(kk < i, kk, i) 'WorksheetFunction.Min(kk, i)
                ll = IIf(ll > i, ll, i) 'WorksheetFunction.Max(ll, i)
                Exit For
            End If
        Next i
    Next j

    TEMP_MATRIX(1, 1 + k) = hh / ii
    TEMP_MATRIX(2, 1 + k) = kk
    TEMP_MATRIX(3, 1 + k) = ll
Next k

PORT_EXPECTED_RECOVERY_PERIOD_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
PORT_EXPECTED_RECOVERY_PERIOD_FUNC = Err.number
End Function
