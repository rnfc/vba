Attribute VB_Name = "FINAN_DERIV_BS_SIMULATION_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : CC_PP_LC_SP_SIMULATION_FUNC
'DESCRIPTION   : Simulation of Optioned Portfolios: Return simulation for Covered
'Call, Protective Put, Long Call and Short Put portfolios and Kernel
'density estimation (non-parametric estimation of the probability
'distribution function).

'LIBRARY       : DERIVATIVES
'GROUP         : PORTFOLIO
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function CC_PP_LC_SP_SIMULATION_FUNC(ByVal STRIKE_VAL As Double, _
ByVal EXPIRATION_VAL As Double, _
ByVal SPOT_VAL As Double, _
ByVal SIGMA_VAL As Double, _
ByVal RATE_VAL As Double, _
ByVal EXPECTED_RETURN_VAL As Double, _
Optional ByVal OPTION_WEIGHT As Double = 0.5, _
Optional ByVal NBINS As Long = 100, _
Optional ByVal nLOOPS As Long = 1000, _
Optional ByVal NTRIALS As Long = 1, _
Optional ByVal VERSION As Integer = 4, _
Optional ByVal OUTPUT As Integer = 1, _
Optional ByVal CND_TYPE As Integer = 0)

'STRIKE_VAL: 100
'EXPIRATION_VAL: 1
'Initial Underlying Price: 100
'Underlying Volatility: 15%
'Riskfree RATE_VAL: 10%
'Expected Return of Underlying: 3%
'% of Portfolio which is optioned: 50%

Dim i As Long
Dim j As Long
Dim DATA_MATRIX As Variant
'% of Portfolio which is optioned

On Error GoTo ERROR_LABEL

'----------------------------------------------------------------------
DATA_MATRIX = MATRIX_RANDOM_NORMAL_FUNC(nLOOPS, NTRIALS, 0, _
              EXPECTED_RETURN_VAL, SIGMA_VAL, 0)
'----------------------------------------------------------------------
Select Case VERSION
'----------------------------------------------------------------------
Case 0 'Long Underlying
'----------------------------------------------------------------------
    
    For j = 1 To NTRIALS
        For i = 1 To nLOOPS
            DATA_MATRIX(i, j) = SPOT_VAL * _
                                Exp(DATA_MATRIX(i, j) * EXPIRATION_VAL)
            DATA_MATRIX(i, j) = Log(DATA_MATRIX(i, j) / SPOT_VAL)
        Next i
    Next j
'----------------------------------------------------------------------
Case 1 'Short Call + Underlying
'----------------------------------------------------------------------
    For j = 1 To NTRIALS
        For i = 1 To nLOOPS
            DATA_MATRIX(i, j) = SPOT_VAL * _
                                Exp(DATA_MATRIX(i, j) * EXPIRATION_VAL)
            
            DATA_MATRIX(i, j) = OPTION_WEIGHT * Log((-MAXIMUM_FUNC( _
                                DATA_MATRIX(i, j) - STRIKE_VAL, 0) + _
                                DATA_MATRIX(i, j) + EUROPEAN_CALL_OPTION_FUNC( _
                                SPOT_VAL, STRIKE_VAL, EXPIRATION_VAL, _
                                RATE_VAL, 0, SIGMA_VAL, _
                                CND_TYPE) * Exp(EXPIRATION_VAL * RATE_VAL)) / _
                                SPOT_VAL) + (1 - OPTION_WEIGHT) * _
                                Log(DATA_MATRIX(i, j) / SPOT_VAL)
        Next i
    Next j
'----------------------------------------------------------------------
Case 2 'Long Put + Underlying
'----------------------------------------------------------------------
    For j = 1 To NTRIALS
        For i = 1 To nLOOPS
            DATA_MATRIX(i, j) = SPOT_VAL * _
                                Exp(DATA_MATRIX(i, j) * EXPIRATION_VAL)
            
            DATA_MATRIX(i, j) = OPTION_WEIGHT * Log((MAXIMUM_FUNC(STRIKE_VAL - _
                                DATA_MATRIX(i, j), 0) + DATA_MATRIX(i, j)) / _
                                (SPOT_VAL + _
                                EUROPEAN_PUT_OPTION_FUNC(SPOT_VAL, _
                                STRIKE_VAL, EXPIRATION_VAL, RATE_VAL, 0, _
                                SIGMA_VAL, CND_TYPE))) + _
                                (1 - OPTION_WEIGHT) * Log(DATA_MATRIX(i, j) / _
                                SPOT_VAL)
        Next i
    Next j
'----------------------------------------------------------------------
Case 3 'Short Put + Underlying
'----------------------------------------------------------------------
    For j = 1 To NTRIALS
        For i = 1 To nLOOPS
            DATA_MATRIX(i, j) = SPOT_VAL * _
                                Exp(DATA_MATRIX(i, j) * EXPIRATION_VAL)
            DATA_MATRIX(i, j) = OPTION_WEIGHT * Log((-MAXIMUM_FUNC(STRIKE_VAL - _
                                DATA_MATRIX(i, j), 0) + DATA_MATRIX(i, j) + _
                                EUROPEAN_PUT_OPTION_FUNC(SPOT_VAL, _
                                STRIKE_VAL, EXPIRATION_VAL, RATE_VAL, 0, _
                                SIGMA_VAL, CND_TYPE) * Exp(EXPIRATION_VAL * _
                                RATE_VAL)) / SPOT_VAL) _
                                + (1 - OPTION_WEIGHT) * Log(DATA_MATRIX(i, j) / _
                                SPOT_VAL)
        Next i
    Next j
'----------------------------------------------------------------------
Case Else 'Long Call + Underlying
'----------------------------------------------------------------------
    For j = 1 To NTRIALS
        For i = 1 To nLOOPS
            DATA_MATRIX(i, j) = SPOT_VAL * _
                                Exp(DATA_MATRIX(i, j) * EXPIRATION_VAL)
            DATA_MATRIX(i, j) = OPTION_WEIGHT * Log((MAXIMUM_FUNC( _
                                DATA_MATRIX(i, j) - STRIKE_VAL, 0) + _
                                DATA_MATRIX(i, j)) / (SPOT_VAL _
                                + EUROPEAN_CALL_OPTION_FUNC( _
                                SPOT_VAL, STRIKE_VAL, _
                                EXPIRATION_VAL, RATE_VAL, 0, _
                                SIGMA_VAL, CND_TYPE))) + _
                                (1 - OPTION_WEIGHT) * _
                                Log(DATA_MATRIX(i, j) / SPOT_VAL)
        Next i
    Next j
'----------------------------------------------------------------------
End Select
'----------------------------------------------------------------------

'----------------------------------------------------------------------
Select Case OUTPUT
'----------------------------------------------------------------------
Case 0
'----------------------------------------------------------------------
    CC_PP_LC_SP_SIMULATION_FUNC = DATA_MATRIX
'----------------------------------------------------------------------
Case 1
'----------------------------------------------------------------------
    CC_PP_LC_SP_SIMULATION_FUNC = DATA_ADVANCED_MOMENTS_FUNC(DATA_MATRIX, 0, 0, 0.05, 0)
'----------------------------------------------------------------------
Case Else
'----------------------------------------------------------------------
    'Probability distributions for Returns of Optioned Portfolios
    CC_PP_LC_SP_SIMULATION_FUNC = FIT_EMPIRICAL_KERNEL_DISTRIBUTION_FUNC(DATA_MATRIX, NBINS) 'X & PDF
'----------------------------------------------------------------------
End Select
'----------------------------------------------------------------------
Exit Function
ERROR_LABEL:
CC_PP_LC_SP_SIMULATION_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : ANTITHETIC_VARIABLES_CALL_OPTION_SIMULATION_FUNC
'DESCRIPTION   : Black Scholes Call price by Monte Carlo with Antithetic variables
'LIBRARY       : DERIVATIVES
'GROUP         : BS_MONTE_CARLO
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function ANTITHETIC_VARIABLES_CALL_OPTION_SIMULATION_FUNC(ByVal SPOT_VAL As Double, _
ByVal STRIKE_VAL As Double, _
ByVal EXPIRATION_VAL As Double, _
ByVal RATE_VAL As Double, _
ByVal SIGMA_VAL As Double, _
Optional ByVal nLOOPS As Long = 1000, _
Optional ByVal OUTPUT As Integer = 0)

Dim i As Long  'paths counter

Dim SPOT1_VAL As Double   ' stock values using z
Dim SPOT2_VAL As Double   ' stock values using -z

Dim NRV_VAL As Double    'standard normal value
Dim OPTION1_VAL As Double    'Black Scholes with z
Dim OPTION2_VAL As Double    'Black Scholes with -z

Dim TEMP_MATRIX As Variant
Dim NRVS_MATRIX As Variant

On Error GoTo ERROR_LABEL

'Beginning the summation
OPTION1_VAL = 0
OPTION2_VAL = 0

If OUTPUT <> 0 Then: ReDim TEMP_MATRIX(1 To nLOOPS, 1 To 2)

NRVS_MATRIX = MATRIX_RANDOM_NORMAL_FUNC(nLOOPS, 1, 0, 0, 1, 0)

For i = 1 To nLOOPS 'Loop over M paths
    NRV_VAL = NRVS_MATRIX(i, 1)
    SPOT1_VAL = SPOT_VAL * Exp((RATE_VAL - ((SIGMA_VAL ^ 2) / 2)) * EXPIRATION_VAL + SIGMA_VAL * (EXPIRATION_VAL ^ 0.5) * NRV_VAL)
    SPOT2_VAL = SPOT_VAL * Exp((RATE_VAL - ((SIGMA_VAL ^ 2) / 2)) * EXPIRATION_VAL - SIGMA_VAL * (EXPIRATION_VAL ^ 0.5) * NRV_VAL)
    OPTION1_VAL = OPTION1_VAL + MAXIMUM_FUNC(SPOT1_VAL - STRIKE_VAL, 0)
    OPTION2_VAL = OPTION2_VAL + MAXIMUM_FUNC(SPOT2_VAL - STRIKE_VAL, 0)
    If OUTPUT <> 0 Then
        TEMP_MATRIX(i, 1) = OPTION1_VAL
        TEMP_MATRIX(i, 2) = OPTION2_VAL
    End If
Next i

Select Case OUTPUT
Case 0
    ANTITHETIC_VARIABLES_CALL_OPTION_SIMULATION_FUNC = Exp(-RATE_VAL * EXPIRATION_VAL) * (OPTION1_VAL + OPTION2_VAL) / (2 * nLOOPS)
Case Else
    ANTITHETIC_VARIABLES_CALL_OPTION_SIMULATION_FUNC = TEMP_MATRIX
End Select

Exit Function
ERROR_LABEL:
ANTITHETIC_VARIABLES_CALL_OPTION_SIMULATION_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : SDE_CALL_OPTION_SIMULATION_FUNC
'DESCRIPTION   : Black Scholes Call price by Monte Carlo with Antithetic
'variables and SDE discretization
'LIBRARY       : DERIVATIVES
'GROUP         : BS_MONTE_CARLO
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************


Function SDE_CALL_OPTION_SIMULATION_FUNC(ByVal SPOT_VAL As Double, _
ByVal STRIKE_VAL As Double, _
ByVal EXPIRATION_VAL As Double, _
ByVal RATE_VAL As Double, _
ByVal SIGMA_VAL As Double, _
Optional ByVal nLOOPS As Long = 1000, _
Optional ByVal nSTEPS As Long = 20, _
Optional ByVal OUTPUT As Integer = 0)

Dim i As Long  'paths counter
Dim j As Long  'time periods counter

Dim SPOT1_VAL As Double   ' stock values using z
Dim SPOT2_VAL As Double   ' stock values using -z

Dim NRV_VAL As Double   'standard normal value
Dim DELTA_VAL As Double  'time step

Dim OPTION1_VAL As Double    'Black Scholes with z
Dim OPTION2_VAL As Double    'Black Scholes with -z

Dim TEMP_MATRIX As Variant
Dim NRVS_MATRIX As Variant

On Error GoTo ERROR_LABEL

'Beginning the summation
OPTION1_VAL = 0
OPTION2_VAL = 0
DELTA_VAL = EXPIRATION_VAL / nSTEPS

If OUTPUT <> 0 Then: ReDim TEMP_MATRIX(1 To nLOOPS, 1 To nSTEPS, 1 To 2)
NRVS_MATRIX = MATRIX_RANDOM_NORMAL_FUNC(nLOOPS, nSTEPS, 0, 0, 1, 0)

For i = 1 To nLOOPS 'Loop over M paths
    SPOT1_VAL = SPOT_VAL
    SPOT2_VAL = SPOT_VAL
    For j = 1 To nSTEPS 'Loop over N time steps
        NRV_VAL = NRVS_MATRIX(i, j)
        SPOT1_VAL = SPOT1_VAL + SPOT1_VAL * (RATE_VAL * DELTA_VAL + SIGMA_VAL * (DELTA_VAL ^ 0.5) * NRV_VAL)
        SPOT2_VAL = SPOT2_VAL + SPOT2_VAL * (RATE_VAL * DELTA_VAL - SIGMA_VAL * (DELTA_VAL ^ 0.5) * NRV_VAL)
        If OUTPUT <> 0 Then
            TEMP_MATRIX(i, j, 1) = SPOT1_VAL
            TEMP_MATRIX(i, j, 2) = SPOT2_VAL
        End If
    Next j
    OPTION1_VAL = OPTION1_VAL + MAXIMUM_FUNC(SPOT1_VAL - STRIKE_VAL, 0)
    OPTION2_VAL = OPTION2_VAL + MAXIMUM_FUNC(SPOT2_VAL - STRIKE_VAL, 0)
Next i

Select Case OUTPUT
Case 0
    SDE_CALL_OPTION_SIMULATION_FUNC = Exp(-RATE_VAL * EXPIRATION_VAL) * (OPTION1_VAL + OPTION2_VAL) / (2 * nLOOPS)
Case Else
    SDE_CALL_OPTION_SIMULATION_FUNC = TEMP_MATRIX
End Select

Exit Function
ERROR_LABEL:
SDE_CALL_OPTION_SIMULATION_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : EUROPEAN_OPTION_SIMULATION_FUNC
'DESCRIPTION   : EUROPEAN OPTION MC SIMULATION
'LIBRARY       : DERIVATIVES
'GROUP         : BS_MONTE_CARLO
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function EUROPEAN_OPTION_SIMULATION_FUNC(ByVal nLOOPS As Long, _
ByVal SPOT_VAL As Double, _
ByVal STRIKE_VAL As Double, _
ByVal EXPIRATION_VAL As Double, _
ByVal RATE_VAL As Double, _
ByVal SIGMA_VAL As Double, _
Optional ByVal OPTION_FLAG As Integer = 1)

'OPTION_FLAG = 1 --> CALL OPTION
'OPTION_FLAG = -1 --> PUT OPTION

Dim i As Long
Dim PAYOFF As Double
Dim NRV_VAL As Double
    
Dim SPOT1_VAL As Double
    
Dim TEMP1_SUM As Double
Dim TEMP2_SUM As Double
    
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL
    
If OPTION_FLAG <> 1 Then OPTION_FLAG = -1
'the default is a call option, unless OPTION_FLAG = -1 to indicate put option

TEMP1_SUM = 0#
TEMP2_SUM = 0#
For i = 1 To nLOOPS
    NRV_VAL = RANDOM_NORMAL_FUNC(0, 1, 0)
    SPOT1_VAL = SPOT_VAL * Exp((RATE_VAL - 0.5 * SIGMA_VAL ^ 2) * EXPIRATION_VAL + SIGMA_VAL * Sqr(EXPIRATION_VAL) * NRV_VAL) 'Geometric Brownian Motion
    PAYOFF = MAXIMUM_FUNC(0#, OPTION_FLAG * SPOT1_VAL - OPTION_FLAG * STRIKE_VAL)
    TEMP1_SUM = TEMP1_SUM + PAYOFF
    TEMP2_SUM = TEMP2_SUM + PAYOFF ^ 2
Next i
      
ReDim TEMP_MATRIX(1 To 2, 1 To 2)
TEMP_MATRIX(1, 1) = "OPTION VALUE"
TEMP_MATRIX(2, 1) = "SE ERROR"
TEMP_MATRIX(1, 2) = Exp(-RATE_VAL * EXPIRATION_VAL) * TEMP1_SUM / nLOOPS
TEMP_MATRIX(2, 2) = (Sqr((TEMP2_SUM - TEMP1_SUM ^ 2 / nLOOPS) / nLOOPS - 1)) / Sqr(nLOOPS)
EUROPEAN_OPTION_SIMULATION_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
EUROPEAN_OPTION_SIMULATION_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : EUROPEAN_CALL_CORPUT_SEQUENCE_SIMULATION_FUNC
'DESCRIPTION   : Corput sequence technique for European Call Option Simulation
'LIBRARY       : DERIVATIVES
'GROUP         : BS_MONTE_CARLO
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function EUROPEAN_CALL_CORPUT_SEQUENCE_SIMULATION_FUNC(ByVal SPOT_VAL As Double, _
ByVal STRIKE_VAL As Double, _
ByVal EXPIRATION_VAL As Double, _
ByVal RATE_VAL As Double, _
ByVal DIVIDEND_VAL As Double, _
ByVal SIGMA_VAL As Double, _
Optional ByVal nLOOPS As Long = 1000)
     
Dim i As Long

Dim C_VAL As Double
Dim NRV_VAL As Double
Dim SPO1_VAL As Double

Dim TEMP_SUM As Double
Dim FACTOR1_VAL As Double
Dim FACTOR2_VAL As Double
          
On Error GoTo ERROR_LABEL
       
FACTOR1_VAL = (RATE_VAL - DIVIDEND_VAL - ((SIGMA_VAL * SIGMA_VAL) / 2)) * EXPIRATION_VAL
FACTOR2_VAL = SIGMA_VAL * Sqr(EXPIRATION_VAL)
TEMP_SUM = 0
'Generating the Quasi-Random vector and the option calculus to feed the simulation
For i = 1 To nLOOPS
    C_VAL = CORPUT_SEQUENCE_NUMBER_FUNC(i, 2)
    NRV_VAL = NORMSINV_FUNC(C_VAL, 0, 1, 1)
    SPO1_VAL = SPOT_VAL * Exp(FACTOR1_VAL + (FACTOR2_VAL * NRV_VAL))
    If SPO1_VAL >= STRIKE_VAL Then: TEMP_SUM = TEMP_SUM + (SPO1_VAL - STRIKE_VAL)
Next i
            
EUROPEAN_CALL_CORPUT_SEQUENCE_SIMULATION_FUNC = (TEMP_SUM / nLOOPS) * Exp(-RATE_VAL * EXPIRATION_VAL)

Exit Function
ERROR_LABEL:
EUROPEAN_CALL_CORPUT_SEQUENCE_SIMULATION_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : BOX_MULLER_CALL_OPTION_SIMULATION_FUNC

'DESCRIPTION   : Uses Box-Muller transformation control variate uses
'(STRIKE_VAL-CONTROL_VAL) antithetic variates
'Option Pricing via Monte Carlo Method using control and antithetic
'variates Vanilla Call

'LIBRARY       : DERIVATIVES
'GROUP         : BS_MONTE_CARLO
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function BOX_MULLER_CALL_OPTION_SIMULATION_FUNC(ByVal SPOT_VAL As Double, _
ByVal STRIKE_VAL As Double, _
ByVal RATE_VAL As Double, _
ByVal SIGMA_VAL As Double, _
ByVal EXPIRATION_VAL As Double, _
Optional ByVal CONTROL_VAL As Double = 5, _
Optional ByVal MIN_LOOPS As Double = 100, _
Optional ByVal MAX_LOOPS As Double = 200, _
Optional ByVal DELTA_LOOPS As Double = 5, _
Optional ByVal RANDOMIZE_INT As Integer = 2, _
Optional ByVal CND_TYPE As Integer = 0)

Dim i As Long
Dim NSIZE As Long

Dim j As Double
Dim NROWS As Double

Dim A_VAL As Double
Dim B_VAL As Double
Dim C_VAL As Double
Dim D_VAL As Double
Dim E_VAL As Double
Dim F_VAL As Double
Dim G_VAL As Double
Dim H_VAL As Double
Dim I_VAL As Double
Dim J_VAL As Double
Dim K_VAL As Double

Dim PI_VAL As Double
Dim TEMP_SUM As Double

Dim VAR_VAL As Double
Dim TEMP_VAL As Double

Dim SPOT1_VAL As Double
Dim SPOT2_VAL As Double

Dim OPTION_VAL As Double

Dim NRV1_VAL As Double
Dim NRV2_VAL As Double

Dim MULT_VAL As Double
Dim FACTOR_VAL As Double

Dim tolerance As Double

On Error GoTo ERROR_LABEL

PI_VAL = 3.14159265358979
NSIZE = CInt((MAX_LOOPS - MIN_LOOPS) / DELTA_LOOPS)

Rnd (-1)
Randomize (RANDOMIZE_INT)
tolerance = 0.000000000000001

'/////////////////////////Randomize [Number]\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'The optional number argument is a Variant or any valid
'numeric expression.

'Randomize uses number to initialize the Rnd function's random-number
'generator, giving it a new seed value. If you omit number, the value
'returned by the system timer is used as the new seed value.

'If Randomize is not used, the Rnd function (with no arguments) uses
'the same number as a seed the first time it is called, and thereafter
'uses the last generated number as a seed value.

'Note: To repeat sequences of random numbers, call Rnd with a negative
'argument immediately before using Randomize with a numeric argument.
'Using Randomize with the same value for number does not repeat the
'previous sequence.
'///////////////////////////////////////////////////////////////////////

'Black-Scholes formula
A_VAL = (Log(SPOT_VAL / STRIKE_VAL) + (RATE_VAL + 0.5 * SIGMA_VAL * SIGMA_VAL) * EXPIRATION_VAL) / (SIGMA_VAL * Sqr(EXPIRATION_VAL))
B_VAL = A_VAL - SIGMA_VAL * Sqr(EXPIRATION_VAL)
C_VAL = CND_FUNC(A_VAL, CND_TYPE)
D_VAL = CND_FUNC(B_VAL, CND_TYPE)

TEMP_VAL = SPOT_VAL * C_VAL - STRIKE_VAL * Exp(-RATE_VAL * EXPIRATION_VAL) * D_VAL

'Black-Scholes formula for the control variate
A_VAL = (Log(SPOT_VAL / (STRIKE_VAL - CONTROL_VAL)) + (RATE_VAL + 0.5 * SIGMA_VAL * SIGMA_VAL) * EXPIRATION_VAL) / (SIGMA_VAL * Sqr(EXPIRATION_VAL))
B_VAL = A_VAL - SIGMA_VAL * Sqr(EXPIRATION_VAL)
C_VAL = CND_FUNC(A_VAL, CND_TYPE)
D_VAL = CND_FUNC(B_VAL, CND_TYPE)
K_VAL = SPOT_VAL * C_VAL - (STRIKE_VAL - CONTROL_VAL) * Exp(-RATE_VAL * EXPIRATION_VAL) * D_VAL

ReDim TEMP_MATRIX(0 To NSIZE, 1 To 4)

TEMP_MATRIX(0, 1) = "TRIALS"
TEMP_MATRIX(0, 2) = "ESTIMATE"
TEMP_MATRIX(0, 3) = "95% CONFIDENCE"
TEMP_MATRIX(0, 4) = "INTERVAL"

For i = 1 To NSIZE

    If i = 1 Then
      j = MIN_LOOPS
    Else: j = MIN_LOOPS + DELTA_LOOPS * (i - 1)
    End If
    If (j < 4) Or (j > 30000) Then: j = 4

    If (j Mod 2 = 1) Then: j = j + 1
    TEMP_MATRIX(i, 1) = j
    NROWS = j / 2

    TEMP_SUM = 0
    VAR_VAL = 0
    C_VAL = SPOT_VAL * Exp((RATE_VAL - 0.5 * SIGMA_VAL * SIGMA_VAL) * EXPIRATION_VAL)
    D_VAL = Sqr(EXPIRATION_VAL)

    For j = 1 To NROWS
        NRV1_VAL = Rnd
        NRV2_VAL = Rnd
        'Box-Muller
        If NRV1_VAL < tolerance Then NRV1_VAL = tolerance 'choose smallest machine number > 0
        
        SPOT1_VAL = Sqr(-2 * Log(NRV1_VAL)) * Cos(2 * PI_VAL * NRV2_VAL)
        SPOT2_VAL = Sqr(-2 * Log(NRV1_VAL)) * Sin(2 * PI_VAL * NRV2_VAL)
        'simulate price in EXPIRATION_VAL
        A_VAL = C_VAL * Exp(SIGMA_VAL * D_VAL * SPOT1_VAL)
        B_VAL = C_VAL * Exp(SIGMA_VAL * D_VAL * SPOT2_VAL)
        'Caculate payoff
        If A_VAL > (STRIKE_VAL - CONTROL_VAL) Then E_VAL = A_VAL - (STRIKE_VAL - CONTROL_VAL) Else E_VAL = 0
        If B_VAL > (STRIKE_VAL - CONTROL_VAL) Then F_VAL = B_VAL - (STRIKE_VAL - CONTROL_VAL) Else F_VAL = 0
        If A_VAL > STRIKE_VAL Then A_VAL = A_VAL - STRIKE_VAL Else A_VAL = 0
        If B_VAL > STRIKE_VAL Then B_VAL = B_VAL - STRIKE_VAL Else B_VAL = 0
     
        TEMP_SUM = TEMP_SUM + (A_VAL - E_VAL) + (B_VAL - F_VAL)
        VAR_VAL = VAR_VAL + (A_VAL - E_VAL) * (A_VAL - E_VAL) + _
                (B_VAL - F_VAL) * (B_VAL - F_VAL)
        
        'Simulate price in EXPIRATION_VAL with antithetic variate
        G_VAL = C_VAL * Exp(SIGMA_VAL * D_VAL * -SPOT1_VAL)
        H_VAL = C_VAL * Exp(SIGMA_VAL * D_VAL * -SPOT2_VAL)
        
        'Calculate payoff
        If G_VAL > (STRIKE_VAL - CONTROL_VAL) Then I_VAL = G_VAL - (STRIKE_VAL - CONTROL_VAL) Else I_VAL = 0
        If H_VAL > (STRIKE_VAL - CONTROL_VAL) Then J_VAL = H_VAL - (STRIKE_VAL - CONTROL_VAL) Else J_VAL = 0
        If G_VAL > STRIKE_VAL Then G_VAL = G_VAL - STRIKE_VAL Else G_VAL = 0
        If H_VAL > STRIKE_VAL Then H_VAL = H_VAL - STRIKE_VAL Else H_VAL = 0

        TEMP_SUM = TEMP_SUM + (G_VAL - I_VAL) + (H_VAL - J_VAL)
        VAR_VAL = VAR_VAL + (G_VAL - I_VAL) * (G_VAL - I_VAL) + (H_VAL - J_VAL) * (H_VAL - J_VAL) + 2 * (G_VAL - I_VAL) * (A_VAL - E_VAL) + 2 * (H_VAL - J_VAL) * (B_VAL - F_VAL)
    Next j

    MULT_VAL = TEMP_SUM / (2 * 2 * NROWS)
    OPTION_VAL = K_VAL + Exp(-RATE_VAL * EXPIRATION_VAL) * MULT_VAL
    FACTOR_VAL = (VAR_VAL / (4 * 2 * NROWS) - MULT_VAL * MULT_VAL) / (2 * NROWS - 1)
    FACTOR_VAL = Sqr(Exp(-RATE_VAL * EXPIRATION_VAL * 2) * FACTOR_VAL)

    TEMP_MATRIX(i, 2) = OPTION_VAL
    TEMP_MATRIX(i, 3) = OPTION_VAL - 1.96 * FACTOR_VAL
    TEMP_MATRIX(i, 4) = OPTION_VAL + 1.96 * FACTOR_VAL
Next i

BOX_MULLER_CALL_OPTION_SIMULATION_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
BOX_MULLER_CALL_OPTION_SIMULATION_FUNC = Err.number
End Function


Function ONE_ASSET_OPTION_MC_FUNC(ByVal SPOT_VAL As Double, _
ByVal SIGMA_VAL As Double, _
ByVal RATE_VAL As Double, _
ByVal STRIKE_VAL As Double, _
ByVal EXPIRATION_VAL As Double, _
Optional ByVal MU_VAL As Double = 0, _
Optional ByVal MATCHED_FLAG As Boolean = True, _
Optional ByVal OPTION_FLAG As Integer = 1, _
Optional ByVal nLOOPS As Long = 10000, _
Optional ByVal OUTPUT As Integer = 1)

Dim i As Long
Dim Z_VAL As Double
Dim NRV_MATRIX As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

If OPTION_FLAG <> 1 Then: OPTION_FLAG = -1 'Call/Put
ReDim TEMP_MATRIX(0 To nLOOPS, 1 To 12)
TEMP_MATRIX(0, 1) = "NRVs"
'MC price = average discounted payoff
TEMP_MATRIX(0, 2) = "BASE SAMPLE: S,T"
TEMP_MATRIX(0, 3) = "BASE SAMPLE: " & IIf(OPTION_FLAG = 1, "CALL", "PUT")
TEMP_MATRIX(0, 4) = "BASE SAMPLE: DISCOUNTED"
'Generate antithetic paths using -N(0,1)
TEMP_MATRIX(0, 5) = "ANTITHETIC PATHS: S,T"
TEMP_MATRIX(0, 6) = "ANTITHETIC PATHS: " & IIf(OPTION_FLAG = 1, "CALL", "PUT")
TEMP_MATRIX(0, 7) = "ANTITHETIC PATHS: AVERAGE DISCOUNTED"
'Generate antithetic paths using -N(0,1) --> Useful for Long maturity Option
TEMP_MATRIX(0, 8) = "IMPORTANCE SAMPLING: S,T"
TEMP_MATRIX(0, 9) = "IMPORTANCE SAMPLING: " & IIf(OPTION_FLAG = 1, "CALL", "PUT")
TEMP_MATRIX(0, 10) = "IMPORTANCE SAMPLING: LIKELIHOOD RATIO"
TEMP_MATRIX(0, 11) = "IMPORTANCE SAMPLING: L: RATIO WEIGHTED"
TEMP_MATRIX(0, 12) = "IMPORTANCE SAMPLING: DISCOUNTED"

Z_VAL = (MU_VAL - RATE_VAL) / SIGMA_VAL
NRV_MATRIX = MULTI_NORMAL_RANDOM_MATRIX_FUNC(0, nLOOPS, 1, 0, 1, True, MATCHED_FLAG, , 0)
For i = 1 To nLOOPS
    TEMP_MATRIX(i, 1) = NRV_MATRIX(i, 1)
    
    'Calculate Stock Price
    TEMP_MATRIX(i, 2) = SPOT_VAL * Exp((RATE_VAL - SIGMA_VAL * SIGMA_VAL / 2) * EXPIRATION_VAL + SIGMA_VAL * Sqr(EXPIRATION_VAL) * TEMP_MATRIX(i, 1))
    TEMP_MATRIX(i, 5) = SPOT_VAL * Exp((RATE_VAL - SIGMA_VAL * SIGMA_VAL / 2) * EXPIRATION_VAL - SIGMA_VAL * Sqr(EXPIRATION_VAL) * TEMP_MATRIX(i, 1))
    TEMP_MATRIX(i, 8) = SPOT_VAL * Exp((MU_VAL - SIGMA_VAL * SIGMA_VAL / 2) * EXPIRATION_VAL + SIGMA_VAL * Sqr(EXPIRATION_VAL) * TEMP_MATRIX(i, 1))
    ' Generate paths using drift mu instead of r
    
    'Calculate Option Payoff
    If OPTION_FLAG = 1 Then
        TEMP_MATRIX(i, 3) = MAXIMUM_FUNC(TEMP_MATRIX(i, 2) - STRIKE_VAL, 0)
        TEMP_MATRIX(i, 6) = MAXIMUM_FUNC(TEMP_MATRIX(i, 5) - STRIKE_VAL, 0)
        TEMP_MATRIX(i, 9) = MAXIMUM_FUNC(TEMP_MATRIX(i, 8) - STRIKE_VAL, 0)
    Else
        TEMP_MATRIX(i, 3) = MAXIMUM_FUNC(TEMP_MATRIX(i, 2) - STRIKE_VAL, 0)
        TEMP_MATRIX(i, 6) = MAXIMUM_FUNC(TEMP_MATRIX(i, 5) - STRIKE_VAL, 0)
        TEMP_MATRIX(i, 9) = MAXIMUM_FUNC(TEMP_MATRIX(i, 8) - STRIKE_VAL, 0)
    End If
    
    'Discount the Payoff
    TEMP_MATRIX(i, 4) = TEMP_MATRIX(i, 3) * Exp(-RATE_VAL * EXPIRATION_VAL)
    TEMP_MATRIX(i, 7) = 0.5 * (TEMP_MATRIX(i, 4) + TEMP_MATRIX(i, 6) * Exp(-RATE_VAL * EXPIRATION_VAL)) 'MC price = average discounted payoff
    TEMP_MATRIX(i, 10) = Exp(-Z_VAL * Sqr(EXPIRATION_VAL) * TEMP_MATRIX(i, 1) - 0.5 * Z_VAL ^ 2 * EXPIRATION_VAL)
    TEMP_MATRIX(i, 11) = TEMP_MATRIX(i, 9) * TEMP_MATRIX(i, 10)
    TEMP_MATRIX(i, 12) = TEMP_MATRIX(i, 11) * Exp(-RATE_VAL * EXPIRATION_VAL)
Next i

Select Case OUTPUT
Case 0
    ONE_ASSET_OPTION_MC_FUNC = TEMP_MATRIX
Case Else
    NRV_MATRIX = DATA_BASIC_MOMENTS_FUNC(TEMP_MATRIX, 0, 0, 0.05, 0)
    ReDim TEMP_MATRIX(1 To 5, 1 To 5)
    TEMP_MATRIX(1, 1) = "METHOD"
    TEMP_MATRIX(2, 1) = "MC"
    TEMP_MATRIX(3, 1) = "AP"
    TEMP_MATRIX(4, 1) = "RW"
    TEMP_MATRIX(5, 1) = "BS" '--> TRUE VALUE
    
    TEMP_MATRIX(1, 2) = "PRICE"
    TEMP_MATRIX(2, 2) = NRV_MATRIX(4, 4)
    TEMP_MATRIX(3, 2) = NRV_MATRIX(7, 4)
    TEMP_MATRIX(4, 2) = NRV_MATRIX(12, 4)
    TEMP_MATRIX(5, 2) = SPOT_VAL * CND_FUNC((Log(SPOT_VAL / STRIKE_VAL) + EXPIRATION_VAL * (RATE_VAL + 0.5 * SIGMA_VAL * SIGMA_VAL)) / (SIGMA_VAL * Sqr(EXPIRATION_VAL)), 0) - STRIKE_VAL * Exp(-RATE_VAL * EXPIRATION_VAL) * CND_FUNC((Log(SPOT_VAL / STRIKE_VAL) + EXPIRATION_VAL * (RATE_VAL - 0.5 * SIGMA_VAL * SIGMA_VAL)) / (SIGMA_VAL * Sqr(EXPIRATION_VAL)), 0)
    
    TEMP_MATRIX(1, 3) = "BS-DIFF"
    TEMP_MATRIX(2, 3) = TEMP_MATRIX(2, 2) / TEMP_MATRIX(5, 2) - 1
    TEMP_MATRIX(3, 3) = TEMP_MATRIX(3, 2) / TEMP_MATRIX(5, 2) - 1
    TEMP_MATRIX(4, 3) = TEMP_MATRIX(4, 2) / TEMP_MATRIX(5, 2) - 1
    TEMP_MATRIX(5, 3) = ""
    
    TEMP_MATRIX(1, 4) = "SE"
    TEMP_MATRIX(2, 4) = NRV_MATRIX(4, 7) / nLOOPS ^ 0.5
    TEMP_MATRIX(3, 4) = NRV_MATRIX(7, 7) / nLOOPS ^ 0.5
    TEMP_MATRIX(4, 4) = NRV_MATRIX(12, 7) / nLOOPS ^ 0.5
    TEMP_MATRIX(5, 4) = ""

    TEMP_MATRIX(1, 5) = "AP:SER^2"
    TEMP_MATRIX(2, 5) = (TEMP_MATRIX(2, 4) / TEMP_MATRIX(3, 4)) ^ 2
    TEMP_MATRIX(3, 5) = (TEMP_MATRIX(3, 4) / TEMP_MATRIX(3, 4)) ^ 2
    TEMP_MATRIX(4, 5) = (TEMP_MATRIX(4, 4) / TEMP_MATRIX(3, 4)) ^ 2
    TEMP_MATRIX(5, 5) = ""

    ONE_ASSET_OPTION_MC_FUNC = TEMP_MATRIX
    
    'SE --> Divide the Stdev of the Discounted values by the nLOOPS ^ 0.5
End Select

Exit Function
ERROR_LABEL:
ONE_ASSET_OPTION_MC_FUNC = Err.number
End Function

Function TWO_ASSETS_OPTION_MC_FUNC(ByVal SPOT1_VAL As Double, _
ByVal SPOT2_VAL As Double, _
ByVal SIGMA1_VAL As Double, _
ByVal SIGMA2_VAL As Double, _
ByVal RHO_VAL As Double, _
ByVal RATE_VAL As Double, _
ByVal STRIKE_VAL As Double, _
ByVal EXPIRATION_VAL As Double, _
Optional ByVal OPTION_FLAG As Integer = 1, _
Optional ByVal nLOOPS As Long = 1000, _
Optional ByVal VERSION As Integer = 0, _
Optional ByVal OUTPUT As Integer = 0)

Dim i As Long
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

If OPTION_FLAG <> 1 Then: OPTION_FLAG = -1 'Call/Put
ReDim TEMP_MATRIX(0 To nLOOPS, 1 To 14)
TEMP_MATRIX(0, 1) = "NRVs1"
TEMP_MATRIX(0, 2) = "NRVs2"
TEMP_MATRIX(0, 3) = "CNRVs" 'Correlated Normal Random Variable
'MC price = average discounted payoff
TEMP_MATRIX(0, 4) = "BASE SAMPLE: S1,T"
TEMP_MATRIX(0, 5) = "BASE SAMPLE: S2,T"
TEMP_MATRIX(0, 6) = "BASE SAMPLE: " & IIf(VERSION = 0, "SUM", "DIFF")
TEMP_MATRIX(0, 7) = "BASE SAMPLE: " & IIf(OPTION_FLAG = 1, "CALL", "PUT")
TEMP_MATRIX(0, 8) = "BASE SAMPLE: DISCOUNTED"
'Generate antithetic paths using -N(0,1)
TEMP_MATRIX(0, 9) = "ANTITHETIC PATHS: S1,T"
TEMP_MATRIX(0, 10) = "ANTITHETIC PATHS: S2,T"
TEMP_MATRIX(0, 11) = "ANTITHETIC PATHS: " & IIf(VERSION = 0, "SUM", "DIFF")
TEMP_MATRIX(0, 12) = "ANTITHETIC PATHS: " & IIf(OPTION_FLAG = 1, "CALL", "PUT")
TEMP_MATRIX(0, 13) = "ANTITHETIC PATHS: DISCOUNTED"
TEMP_MATRIX(0, 14) = "ANTITHETIC PATHS: AVERAGE DISCOUNTED"

For i = 1 To nLOOPS
    TEMP_MATRIX(i, 1) = RANDOM_NORMAL_FUNC(0, 1, 0)
    TEMP_MATRIX(i, 2) = RANDOM_NORMAL_FUNC(0, 1, 0)
    TEMP_MATRIX(i, 3) = (RHO_VAL * TEMP_MATRIX(i, 1)) + (Sqr(1 - RHO_VAL * RHO_VAL) * TEMP_MATRIX(i, 2))
    
    TEMP_MATRIX(i, 4) = SPOT1_VAL * Exp((RATE_VAL - SIGMA1_VAL * SIGMA1_VAL / 2) * EXPIRATION_VAL + SIGMA1_VAL * Sqr(EXPIRATION_VAL) * TEMP_MATRIX(i, 1))
    TEMP_MATRIX(i, 5) = SPOT2_VAL * Exp((RATE_VAL - SIGMA2_VAL * SIGMA2_VAL / 2) * EXPIRATION_VAL + SIGMA2_VAL * Sqr(EXPIRATION_VAL) * TEMP_MATRIX(i, 3))
    
    TEMP_MATRIX(i, 9) = SPOT1_VAL * Exp((RATE_VAL - SIGMA1_VAL * SIGMA1_VAL / 2) * EXPIRATION_VAL - SIGMA1_VAL * Sqr(EXPIRATION_VAL) * TEMP_MATRIX(i, 1))
    TEMP_MATRIX(i, 10) = SPOT1_VAL * Exp((RATE_VAL - SIGMA1_VAL * SIGMA1_VAL / 2) * EXPIRATION_VAL - SIGMA1_VAL * Sqr(EXPIRATION_VAL) * TEMP_MATRIX(i, 3))
    
    If VERSION = 0 Then
        TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 4) + TEMP_MATRIX(i, 5)
        TEMP_MATRIX(i, 11) = TEMP_MATRIX(i, 9) + TEMP_MATRIX(i, 10)
    Else
        TEMP_MATRIX(i, 6) = Abs(TEMP_MATRIX(i, 5) - TEMP_MATRIX(i, 4))
        TEMP_MATRIX(i, 11) = Abs(TEMP_MATRIX(i, 9) - TEMP_MATRIX(i, 10))
    End If
    If OPTION_FLAG = 1 Then
        TEMP_MATRIX(i, 7) = MAXIMUM_FUNC(TEMP_MATRIX(i, 6) - STRIKE_VAL, 0)
        TEMP_MATRIX(i, 12) = MAXIMUM_FUNC(TEMP_MATRIX(i, 11) - STRIKE_VAL, 0)
    Else
        TEMP_MATRIX(i, 7) = MAXIMUM_FUNC(STRIKE_VAL - TEMP_MATRIX(i, 6), 0)
        TEMP_MATRIX(i, 12) = MAXIMUM_FUNC(STRIKE_VAL - TEMP_MATRIX(i, 11), 0)
    End If
    TEMP_MATRIX(i, 8) = TEMP_MATRIX(i, 7) * Exp(-RATE_VAL * EXPIRATION_VAL)
    TEMP_MATRIX(i, 13) = TEMP_MATRIX(i, 12) * Exp(-RATE_VAL * EXPIRATION_VAL)
    TEMP_MATRIX(i, 14) = (TEMP_MATRIX(i, 8) + TEMP_MATRIX(i, 13)) * 0.5
Next i

Select Case OUTPUT
Case 0
    TWO_ASSETS_OPTION_MC_FUNC = TEMP_MATRIX
Case Else
    TWO_ASSETS_OPTION_MC_FUNC = DATA_BASIC_MOMENTS_FUNC(TEMP_MATRIX, 0, 0, 0.05, 1)
    'SE --> Divide the Stdev of the Discounted values by the nLOOPS ^ 0.5
End Select

Exit Function
ERROR_LABEL:
TWO_ASSETS_OPTION_MC_FUNC = Err.number
End Function

