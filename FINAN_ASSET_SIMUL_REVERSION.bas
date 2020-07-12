Attribute VB_Name = "FINAN_ASSET_SIMUL_REVERSION"

'----------------------------------------------------------------------------------
'----------------------------------------------------------------------------------
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'----------------------------------------------------------------------------------
'----------------------------------------------------------------------------------

'************************************************************************************
'************************************************************************************
'FUNCTION      : ASSET_MEAN_REVERSION_SAMPLING_FUNC

'DESCRIPTION   : Mean Revertion Function: Short-run and long-run impacts of
'a change in an exogenous variable
'Once upon a time there was once an article where the author, in explaining
'Reversion to the Mean (RTM), said something like:

'The mathematical principle of reversion (or regression) to the mean states that
'"the greater the deviation of a random variate from its mean, the greater the
'probability that the next measured variate will deviate less far."

'The classic example is a series of coin tosses. If a coin comes up Heads 90 times out
'of the first 100 tosses, look for Tails to make a comeback over the next 100.

'LIBRARY       : FINAN_ASSET
'GROUP         : SIMULATION_REVERSION
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function ASSET_MEAN_REVERSION_SAMPLING_FUNC(ByVal DESIRED_LEVEL As Double, _
ByVal SPOT_VALUE As Double, _
ByVal ALPHA As Double, _
ByVal START_DATE As Date, _
ByVal END_DATE As Date, _
ByVal DELTA_TIME As Double, _
Optional ByVal HOLIDAYS_RNG As Variant)

'ALPHA         --> ADJUSTMENT PARAMETER
'DESIRED_LEVEL --> DESIRED_LEVEL OF SHORT RATE

Dim i As Long
Dim NROWS As Long
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

NROWS = NETWORKDAYS_FUNC(START_DATE, END_DATE, HOLIDAYS_RNG)

ReDim TEMP_MATRIX(0 To NROWS + 1, 1 To 3)

TEMP_MATRIX(0, 1) = "PERIOD"
TEMP_MATRIX(0, 2) = "SPOT"
TEMP_MATRIX(0, 3) = "SPOT*"

TEMP_MATRIX(1, 1) = START_DATE
TEMP_MATRIX(1, 2) = SPOT_VALUE
TEMP_MATRIX(1, 3) = DESIRED_LEVEL

i = 2
Do While i <= NROWS + 1
    TEMP_MATRIX(i, 1) = WORKDAY2_FUNC(TEMP_MATRIX(i - 1, 1), DELTA_TIME, HOLIDAYS_RNG)
    TEMP_MATRIX(i, 2) = TEMP_MATRIX(i - 1, 2) + ALPHA * (DESIRED_LEVEL - TEMP_MATRIX(i - 1, 2))
    TEMP_MATRIX(i, 3) = DESIRED_LEVEL
    i = i + 1
Loop

ASSET_MEAN_REVERSION_SAMPLING_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
ASSET_MEAN_REVERSION_SAMPLING_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : ASSET_MEAN_REVERSION_SIMULATION_FUNC
'DESCRIPTION   : Geometric Brownian Motion Simulation Function
'LIBRARY       : FINAN_ASSET
'GROUP         : SIMULATION_REVERSION
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function ASSET_MEAN_REVERSION_SIMULATION_FUNC(ByVal SPOT As Double, _
ByVal MEAN_VAL As Double, _
ByVal SIGMA_VAL As Double, _
ByVal START_DATE As Date, _
ByVal END_DATE As Date, _
Optional ByVal nLOOPS As Long = 30, _
Optional ByRef HOLIDAYS_RNG As Variant, _
Optional ByVal COUNT_BASIS As Double = 252, _
Optional ByVal OUTPUT As Integer = 1, _
Optional ByVal RANDOM_FLAG As Boolean = True)

Dim i As Long
Dim j As Long
Dim PERIODS As Long

Dim DELTA_TIME As Double

Dim TEMP_MATRIX As Variant
Dim TEMP_VECTOR As Variant

Dim NORMAL_RANDOM_MATRIX As Variant

On Error GoTo ERROR_LABEL

PERIODS = NETWORKDAYS_FUNC(START_DATE, END_DATE, HOLIDAYS_RNG)
DELTA_TIME = 1 / COUNT_BASIS

ReDim TEMP_MATRIX(0 To PERIODS + 1, 1 To nLOOPS + 2)
ReDim TEMP_VECTOR(1 To nLOOPS, 1 To 1)

TEMP_MATRIX(0, 1) = "PERIOD"
TEMP_MATRIX(0, 2) = "DATE"

TEMP_MATRIX(1, 1) = 0
TEMP_MATRIX(1, 2) = START_DATE

For i = 1 To PERIODS
    TEMP_MATRIX(i + 1, 1) = TEMP_MATRIX(i, 1) + DELTA_TIME
    TEMP_MATRIX(i + 1, 2) = WORKDAY2_FUNC(TEMP_MATRIX(i, 2), 1, HOLIDAYS_RNG)
Next i

If RANDOM_FLAG = True Then: Randomize
NORMAL_RANDOM_MATRIX = MULTI_NORMAL_RANDOM_MATRIX_FUNC(0, PERIODS + 2, nLOOPS, 0, 1, RANDOM_FLAG, True, , 0)

For j = 1 To nLOOPS
    TEMP_MATRIX(0, j + 2) = "TRIAL: " & Format(j, "0")
    TEMP_MATRIX(1, j + 2) = SPOT
    
    For i = 1 To PERIODS
        TEMP_MATRIX(i + 1, j + 2) = _
            ASSET_MEAN_REVERSION_SPOT_FUNC(TEMP_MATRIX(i, j + 2), _
            MEAN_VAL, SIGMA_VAL, NORMAL_RANDOM_MATRIX(i + 1, j), DELTA_TIME)
    Next i
    
    TEMP_VECTOR(j, 1) = TEMP_MATRIX(PERIODS + 1, j + 2)
Next j

Select Case OUTPUT
Case 0
    ASSET_MEAN_REVERSION_SIMULATION_FUNC = TEMP_VECTOR
Case Else
    ASSET_MEAN_REVERSION_SIMULATION_FUNC = TEMP_MATRIX
End Select

Exit Function
ERROR_LABEL:
ASSET_MEAN_REVERSION_SIMULATION_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : ASSET_MEAN_REVERSION_SPOT_FUNC
'DESCRIPTION   : Geometric Brownian Motion Discrete Function
'LIBRARY       : FINAN_ASSET
'GROUP         : SIMULATION_REVERSION
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

'Source: Hull, John C., Options, Futures & Other Derivatives. Fourth edition
'(2000). Prentice-Hall. P. 220

Private Function ASSET_MEAN_REVERSION_SPOT_FUNC(ByVal PREVIOUS_SPOT_VAL As Double, _
ByVal MEAN_VAL As Double, _
ByVal SIGMA_VAL As Double, _
ByVal NORMAL_RANDOM_VAL As Double, _
ByVal DELTA_TIME As Double)

'--------------------------------------------------------------------------------
'Change in the asset price, is in a small interval of time.
'--------------------------------------------------------------------------------
'MEAN_VAL:  Expected Return Annualized
'SIGMA_VAL: Sigma Annualized
'NORMAL_RANDOM_VAL: Is a random drawing from a standardized normal
'                   distribution, F(0,1)
'DELTA_TIME = TENOR / PERIODS
'--------------------------------------------------------------------------------

 On Error GoTo ERROR_LABEL

ASSET_MEAN_REVERSION_SPOT_FUNC = _
    PREVIOUS_SPOT_VAL * Exp((MEAN_VAL - 0.5 * SIGMA_VAL ^ 2) * _
    DELTA_TIME + SIGMA_VAL * (DELTA_TIME) ^ 0.5 * NORMAL_RANDOM_VAL)

Exit Function
ERROR_LABEL:
ASSET_MEAN_REVERSION_SPOT_FUNC = Err.number
End Function


'Spot Path through Geometric Brownian Motion (Returns 6 Draws at once)

Function ASSET_SPOT_PATH_SAMPLING_FUNC(ByVal SPOT As Double, _
ByVal TENOR As Double, _
ByVal Mean As Double, _
ByVal SIGMA As Double, _
Optional ByVal nPATHS As Long = 6, _
Optional ByVal COUNT_BASIS As Double = 252)

'The variable  dSPOT is the change in the asset price, S, in a
'small interval of time (dT are independent).
 
'MEAN is the expected rate of return per unit of time (usually per year)
'SIGMA is the volatility of the asset price (s and m are assumed contant)
'RND_NO is a random drawing from a standardized normal distribution, F(0,1)

'MEAN --> ANNUALIZED
'SIGMA --> ANNUALIZED

Dim i As Long
Dim j As Long
Dim PERIODS As Long

Dim DELTA_TIME As Double
Dim START_PERIOD As Double

Dim TEMP_MATRIX As Variant
Dim NORMAL_RANDOM_MATRIX As Variant

On Error GoTo ERROR_LABEL

PERIODS = TENOR * COUNT_BASIS
DELTA_TIME = TENOR / PERIODS
START_PERIOD = 0

ReDim TEMP_MATRIX(0 To PERIODS + 1, 1 To 2 + nPATHS)

TEMP_MATRIX(0, 1) = "PERIOD"
TEMP_MATRIX(0, 2) = "TIME"
j = 3
For i = 1 To nPATHS
    TEMP_MATRIX(0, j) = "S" & CStr(i) & " + DS"
    TEMP_MATRIX(1, j) = SPOT
    j = j + 1
Next i

TEMP_MATRIX(1, 1) = START_PERIOD
TEMP_MATRIX(1, 2) = START_PERIOD * DELTA_TIME

NORMAL_RANDOM_MATRIX = MATRIX_RANDOM_NORMAL_FUNC(PERIODS + 1, nPATHS, 0, 0, 1, 0)

For i = 1 To PERIODS
    TEMP_MATRIX(i + 1, 1) = TEMP_MATRIX(i, 1) + 1 'PERIOD
    TEMP_MATRIX(i + 1, 2) = TEMP_MATRIX(i, 2) + DELTA_TIME 'TIME
    For j = 1 To nPATHS
        TEMP_MATRIX(i + 1, 2 + j) = _
            Mean * TEMP_MATRIX(i, 2 + j) * (DELTA_TIME) + SIGMA * _
            TEMP_MATRIX(i, 2 + j) * NORMAL_RANDOM_MATRIX(i, j) * (DELTA_TIME) ^ 0.5 _
            + TEMP_MATRIX(i, 2 + j) 'S + dS --> j DRAW
    Next j
Next i

ASSET_SPOT_PATH_SAMPLING_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
ASSET_SPOT_PATH_SAMPLING_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : ASSET_SPOT_JUMP_SAMPLING_FUNC
'DESCRIPTION   : Geometric Brownian Motion with Jump Term
'LIBRARY       : FINAN_ASSET
'GROUP         : SIMULATION_REVERSION
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

'REFERENCE: Hull, John C., Options, Futures & Other Derivatives.
'Fourth edition (2000). Prentice-Hall. Brownian motion: p. 226,
'Jump Diffusion (Merton): p. 446

Function ASSET_SPOT_JUMP_SAMPLING_FUNC(ByVal SPOT As Double, _
ByVal TENOR As Double, _
ByVal SIGMA As Double, _
ByVal DRIFT As Double, _
ByVal LAMBDA As Double, _
ByVal KAPPA As Double, _
Optional ByVal SIGMA_GAMMA As Double = 0.5, _
Optional ByVal START_PERIOD As Long = 0, _
Optional ByVal COUNT_BASIS As Double = 252)

'The variable dSPOT is the change in the stock price, S, in a small interval
'of time (dT are independent).

'-----------------------------------Parameters-------------------------------------

'DRIFT: is the expected rate of return per unit of time (usually per year)

'SIGMA: is the volatility of the asset price (s and m are assumed contant)

'NORM_RND_NO: is a random drawing from a standardized normal distribution, F(0,1)

'KAPPA: Average jump size measured as a proportional increase in asset price
'as % of previous stock price

'LAMBDA: Rate at which jumps happen per time unit (year) --> number of jumps per year
'also called INTENSITY of Poisson process

'SIGMA_GAMMA: Standard deviation log of J(Sigma')

'----------------------------------------------------------------------------------

Dim i As Long
Dim PERIODS As Long

Dim DELTA_TIME As Double
Dim DELTA_PRICE As Double

Dim GAMMA_VAL As Double
Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double
Dim CTEMP_VAL As Double

Dim NORMAL_RAND_VAL As Double
Dim UNIFORM_RAND_VAL As Double

Dim TEMP_MATRIX As Variant
Dim NORMAL_RANDOM_MATRIX As Variant
Dim UNIFORM_RANDOM_MATRIX As Variant

On Error GoTo ERROR_LABEL

PERIODS = TENOR * COUNT_BASIS
DELTA_TIME = TENOR / PERIODS
DELTA_PRICE = 1 + KAPPA ''% asset price after jump (J)
'as % of previous asset price

GAMMA_VAL = Log(DELTA_PRICE) 'Drift of ln(J)

ReDim TEMP_MATRIX(0 To PERIODS + 1, 1 To 7) '8)

i = 0
TEMP_MATRIX(i, 1) = "PERIOD"
TEMP_MATRIX(i, 2) = "TIME"
TEMP_MATRIX(i, 3) = "PURE_BROWNIAN_MOTION"
TEMP_MATRIX(i, 4) = "PURE_BROWNIAN_MOTION_FIXED_JUMP"
TEMP_MATRIX(i, 5) = "FIXED_JUMP_SIZE"
TEMP_MATRIX(i, 6) = "PURE_BROWNIAN_MOTION_DYNAMIC_LOG_JUMP"
TEMP_MATRIX(i, 7) = "DYNAMIC_LOG_JUMP_SIZE"

i = 1
TEMP_MATRIX(i, 1) = START_PERIOD
TEMP_MATRIX(i, 2) = TEMP_MATRIX(1, 1) * DELTA_TIME
TEMP_MATRIX(i, 3) = SPOT
TEMP_MATRIX(i, 4) = SPOT
TEMP_MATRIX(i, 5) = 0
TEMP_MATRIX(i, 6) = SPOT
TEMP_MATRIX(i, 7) = 0
'TEMP_MATRIX(i, 8) = 0

NORMAL_RANDOM_MATRIX = MATRIX_RANDOM_NORMAL_FUNC(PERIODS + 1, 1, 0, 0, 1, 0)
'Normal variate for jump size process
UNIFORM_RANDOM_MATRIX = MATRIX_RANDOM_UNIFORM_FUNC(PERIODS, 1, 1, 0)
' Uniform random number for jump process

'-------------------------------------------------------------------------------
For i = 1 To PERIODS
'-------------------------------------------------------------------------------
    ATEMP_VAL = 0: BTEMP_VAL = 0: CTEMP_VAL = 0
    NORMAL_RAND_VAL = NORMAL_RANDOM_MATRIX(i + 1, 1) 'NORMD_RND_ARR(1,1) is for --> i = 0
    UNIFORM_RAND_VAL = UNIFORM_RANDOM_MATRIX(i, 1) 'UNIFORM_RAND_VAL(1,1) is for --> i = 0
    TEMP_MATRIX(i + 1, 1) = TEMP_MATRIX(i, 1) + 1  'PERIOD
    TEMP_MATRIX(i + 1, 2) = TEMP_MATRIX(i, 2) + DELTA_TIME 'TIME
'------------------------------FIXED JUMP SIZE----------------------------------
    TEMP_MATRIX(i + 1, 3) = TEMP_MATRIX(i, 3) * (1 + DRIFT * (DELTA_TIME) + SIGMA * NORMAL_RAND_VAL * (DELTA_TIME) ^ 0.5) 'FIXED JUMP SIZE: Pure Brownian Motion
    If (DELTA_TIME * LAMBDA) > UNIFORM_RAND_VAL Then: ATEMP_VAL = KAPPA
    TEMP_MATRIX(i + 1, 4) = TEMP_MATRIX(i, 4) * (1 + DRIFT * (DELTA_TIME) + SIGMA * NORMAL_RAND_VAL * (DELTA_TIME) ^ 0.5 + ATEMP_VAL) 'FIXED JUMP SIZE: Brownian motion with fixed jumps
    TEMP_MATRIX(i + 1, 5) = TEMP_MATRIX(i, 4) * ATEMP_VAL + TEMP_MATRIX(i, 5)  'JUMP
'----------------------------LOGNORMALLY DISTRIBUTED JUMP SIZE------------------
'Brownian motion with stochastic jumps (ln J normally distributed)
'-------------------------------------------------------------------------------
    If ATEMP_VAL <> 0 Then
        BTEMP_VAL = Exp(GAMMA_VAL + NORMAL_RANDOM_MATRIX(i, 1) * SIGMA_GAMMA) - 1
        CTEMP_VAL = (TEMP_MATRIX(i, 6) * BTEMP_VAL)
    End If
    TEMP_MATRIX(i + 1, 6) = TEMP_MATRIX(i, 6) * (1 + DRIFT * (DELTA_TIME) + SIGMA * NORMAL_RAND_VAL * (DELTA_TIME) ^ 0.5 + BTEMP_VAL) 'FIXED JUMP SIZE: Pure Brownian Motion
    TEMP_MATRIX(i + 1, 7) = CTEMP_VAL + TEMP_MATRIX(i, 7)
    'FIXED JUMP SIZE: Brownian motion with fixed jumps.
    'TEMP_MATRIX(i + 1, 8) = Exp(GAMMA_VAL + NORMAL_RANDOM_MATRIX(i, 1) * SIGMA_GAMMA)
'-------------------------------------------------------------------------------
Next i
'-------------------------------------------------------------------------------

ASSET_SPOT_JUMP_SAMPLING_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
ASSET_SPOT_JUMP_SAMPLING_FUNC = Err.number
End Function
