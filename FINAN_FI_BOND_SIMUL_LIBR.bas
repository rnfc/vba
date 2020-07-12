Attribute VB_Name = "FINAN_FI_BOND_SIMUL_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
                    
'**********************************************************************************
'**********************************************************************************
'FUNCTION      : SHORT_RATE_SIMULATION_FUNC
'DESCRIPTION   : Interest Rate One-Factor Equilibrium Simulation Models:
'"Simulation One-Factor Models" shows discrete versions of
'CIR and VAISCEK models so one gets a feel for their mean-reverting
'nature of the stochastic processes.
'LIBRARY       : FIXED_INCOME
'GROUP         : SIMULATION
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'**********************************************************************************
'**********************************************************************************

Function SHORT_RATE_SIMULATION_FUNC(ByVal START_DATE As Date, _
ByVal END_DATE As Date, _
ByVal nLOOPS As Long, _
ByVal SHORT_RATE As Double, _
ByVal EQUILIBRIUM As Double, _
ByVal PULL_BACK As Double, _
ByVal SIGMA As Double, _
Optional ByVal ZERO_STR_NAME As String = "VASICEK_DISCR_FUNC", _
Optional ByRef HOLIDAYS_RNG As Variant, _
Optional ByVal COUNT_BASIS As Double = 252, _
Optional ByVal OUTPUT As Integer = 0, _
Optional ByVal RANDOM_FLAG As Boolean = True)

'Interest Rate One-Factor Equilibrium Models:
'-----> Stochastic process for short-term interest rate


'--> Source: Hull, John C., Options, Futures & Other Derivatives. Fourth edition
'(2000). Prentice-Hall. P. 567.

'--> Vasicek, O. 1977 "An Equilibrium Characterization of the term structure."
'Journal of Financial Economics 5: 177-188.

'--> Cox, Ingersoll, and Ross. "A Theory of the Term Structure of Interest Rates".
'Econometrica, 53 (1985). 385-407.


Dim i As Long
Dim j As Long
Dim PERIODS As Long

Dim DELTA_TENOR As Double

Dim TEMP_MATRIX As Variant
Dim TEMP_VECTOR As Variant
Dim NORMAL_RANDOM_MATRIX As Variant

On Error GoTo ERROR_LABEL

PERIODS = NETWORKDAYS_FUNC(START_DATE, END_DATE, HOLIDAYS_RNG)
DELTA_TENOR = 1 / COUNT_BASIS

ReDim TEMP_MATRIX(0 To PERIODS + 1, 1 To nLOOPS + 2)
ReDim TEMP_VECTOR(1 To nLOOPS, 1 To 1)

TEMP_MATRIX(0, 1) = ("TN")
TEMP_MATRIX(0, 2) = ("DATE")

TEMP_MATRIX(1, 1) = 0
TEMP_MATRIX(1, 2) = START_DATE

For i = 1 To PERIODS
    TEMP_MATRIX(i + 1, 1) = TEMP_MATRIX(i, 1) + DELTA_TENOR
    TEMP_MATRIX(i + 1, 2) = WORKDAY2_FUNC(TEMP_MATRIX(i, 2), 1, HOLIDAYS_RNG)
Next i

If RANDOM_FLAG = True Then: Randomize
NORMAL_RANDOM_MATRIX = MULTI_NORMAL_RANDOM_MATRIX_FUNC(0, PERIODS + 2, nLOOPS, 0, 1, RANDOM_FLAG, True, , 0)

For j = 1 To nLOOPS
    TEMP_MATRIX(0, j + 2) = "TRIAL: " & Format(j, "0")
    TEMP_MATRIX(1, j + 2) = SHORT_RATE
    For i = 1 To PERIODS
        TEMP_MATRIX(i + 1, j + 2) = Excel.Application.Run(ZERO_STR_NAME, DELTA_TENOR, TEMP_MATRIX(i, j + 2), EQUILIBRIUM, PULL_BACK, SIGMA, NORMAL_RANDOM_MATRIX(i + 1, j))
    Next i
    TEMP_VECTOR(j, 1) = TEMP_MATRIX(PERIODS + 1, j + 2)
Next j

Select Case OUTPUT
Case 0
    SHORT_RATE_SIMULATION_FUNC = TEMP_VECTOR
Case 1
    SHORT_RATE_SIMULATION_FUNC = TEMP_MATRIX
End Select

Exit Function
ERROR_LABEL:
SHORT_RATE_SIMULATION_FUNC = Err.number
End Function
