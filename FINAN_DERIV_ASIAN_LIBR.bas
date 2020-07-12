Attribute VB_Name = "FINAN_DERIV_ASIAN_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_OPTION_ASIAN_SIMULATION_FUNC
'DESCRIPTION   : ASIAN_CALL_SIMUL_FUNC
'LIBRARY       : DERIVATIVES
'GROUP         : ASIAN
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function CALL_OPTION_ASIAN_SIMULATION_FUNC(ByVal nLOOPS As Long, _
ByVal DELTA As Double, _
ByVal SPOT As Double, _
ByVal STRIKE As Double, _
ByVal EXPIRATION As Double, _
ByVal RATE As Double, _
ByVal SIGMA As Double, _
Optional ByVal OUTPUT As Integer = 0)

' nLOOPS = number of Monte Carlo replications
' DELTA = partition of time DELTA

Dim i As Long
Dim j As Long

Dim NRV_VAL As Double
Dim MULT_VAL As Double
Dim TEMP_SUM As Double
Dim SPOT1_VAL As Double
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

ReDim TEMP_MATRIX(0 To nLOOPS, 1 To 3)

TEMP_MATRIX(0, 1) = "LOOP"
TEMP_MATRIX(0, 2) = "PRICE"
TEMP_MATRIX(0, 3) = "CUMULATIVE"

TEMP_SUM = 0

For i = 1 To nLOOPS
    For j = 1 To DELTA
        NRV_VAL = RANDOM_NORMAL_FUNC(0, 1, 0)
        SPOT1_VAL = ((RATE - 0.5 * SIGMA * SIGMA) * EXPIRATION / DELTA) + SIGMA * Sqr(EXPIRATION / DELTA) * NRV_VAL
        MULT_VAL = 1 + Exp(SPOT1_VAL) * MULT_VAL
    Next j
    
    TEMP_SUM = TEMP_SUM + MAXIMUM_FUNC(0, SPOT * MULT_VAL / (DELTA + 1) - i)

    TEMP_MATRIX(i, 1) = i
    TEMP_MATRIX(i, 2) = SPOT1_VAL
    TEMP_MATRIX(i, 3) = TEMP_SUM
Next i


TEMP_SUM = TEMP_SUM / nLOOPS

Select Case OUTPUT
Case 0
    CALL_OPTION_ASIAN_SIMULATION_FUNC = Exp(-RATE * EXPIRATION) * TEMP_SUM
Case Else
    CALL_OPTION_ASIAN_SIMULATION_FUNC = TEMP_MATRIX
End Select


Exit Function
ERROR_LABEL:
CALL_OPTION_ASIAN_SIMULATION_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : ASIAN_SPREAD_SIMULATION_FUNC
'DESCRIPTION   : ASIAN_SPREAD_SIMULATION_FUNC
'LIBRARY       : DERIVATIVES
'GROUP         : ASIAN
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function ASIAN_SPREAD_SIMULATION_FUNC(ByVal nLOOPS As Long, _
ByVal SPOT As Double, _
ByVal STRIKE As Double, _
ByVal START_TENOR As Double, _
ByVal ORIGINAL_TENOR As Double, _
ByVal RATE As Double, _
ByVal CARRY_COST As Double, _
ByVal SIGMA As Double, _
ByVal STEPS As Long, _
Optional ByVal OPTION_FLAG As Integer = 1, _
Optional ByVal OUTPUT As Integer = 0)

'START_TENOR: Time to start of average period
'ORIGINAL_TENOR: Original time to maturity
'STEPS: Number of time steps for simulation
'nLOOPS: Number of simulations

'Extensive testing has shown that the unconditional
'Asian Monte Carlo simulation produces results equal to
'the Turnbull - Wakeman approximation within the random
'variation inherent in the simulation.

'Our technique therefore uses the Turnbull-Wakeman
'result as the basis for our valuation and only uses the
'simulation to estimate the adjustment necessary to
'account for the conditions placed on the Asian option.
'In this way, the variability inherent in simulation is
'confined to a small percentage of the total result.


'For the Monte Carlo simulation, T should always equal T2.
'SAV, the average price from T to T2, is not used
'in the simulation, even though Turnbull-Wakeman
'does have this capability.

'Tau, the time to the start of the averaging period, can
'be used by both the simulation and Turnbull-Wakeman.
'Note, however, that Turnbull-Wakeman also requires that
'T = T2, if Tau is not equal to 0.

'The length of the averaging period is set by the difference between
'the Time to the start of the averaging period and the time to maturity.
'Thus, when the option has a 5 year maturity and a 3-year averaging
'period, one must set the time to the start of the averaging period
'equal to 2.
    
Dim i As Long
Dim j As Long
Dim k As Long

Dim TEMP1_SUM As Double
Dim TEMP2_SUM As Double
Dim TEMP3_SUM As Double

Dim DRIFT_VAL As Double
Dim SIGMA_VAL As Double
Dim MEAN_VAL As Double

Dim NRV_VAL As Double
Dim DT_VAL As Double

On Error GoTo ERROR_LABEL

If OPTION_FLAG <> 1 Then: OPTION_FLAG = -1 'Put
DT_VAL = ORIGINAL_TENOR / STEPS
DRIFT_VAL = (CARRY_COST - SIGMA ^ 2 / 2) * DT_VAL
SIGMA_VAL = SIGMA * Sqr(DT_VAL)

For i = 1 To nLOOPS
    MEAN_VAL = 0
    k = 0
    TEMP3_SUM = SPOT
    For j = 1 To STEPS
        NRV_VAL = RANDOM_NORMAL_FUNC(0, 1, 0)
        TEMP3_SUM = TEMP3_SUM * Exp(DRIFT_VAL + SIGMA_VAL * NRV_VAL)
        If (j * DT_VAL) > START_TENOR Then
            MEAN_VAL = MEAN_VAL + TEMP3_SUM
            k = k + 1
        End If
    Next j
        
    MEAN_VAL = MEAN_VAL / k
    TEMP1_SUM = TEMP1_SUM + MAXIMUM_FUNC(OPTION_FLAG * (MEAN_VAL - STRIKE), 0)
    If TEMP3_SUM < STRIKE Then: MEAN_VAL = 0
    TEMP2_SUM = TEMP2_SUM + MAXIMUM_FUNC(OPTION_FLAG * (MEAN_VAL - STRIKE), 0)
Next i

Select Case OUTPUT
Case 0 ' Unconditional Asian Simulation
    ASIAN_SPREAD_SIMULATION_FUNC = Exp(-RATE * ORIGINAL_TENOR) * (TEMP1_SUM / nLOOPS)
Case Else ' Conditional Asian Simulation
    ASIAN_SPREAD_SIMULATION_FUNC = Exp(-RATE * ORIGINAL_TENOR) * (TEMP2_SUM / nLOOPS)
End Select
   
Exit Function
ERROR_LABEL:
ASIAN_SPREAD_SIMULATION_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : BINOMIAL_ASIAN_OPTION_PRICE_FUNC

'DESCRIPTION   : Asian option price using binomial tree
'This routine calculates price of an asian option using binomial tree method
'as suggested in paper: Hull, J., and A. White, "Efficient Procedures for
'Valuing European and American Path-Dependent Options," Journal of Derivatives,
'Volume 1, pp. 21-31. It compares results as mentioned in Exhibit 4

'LIBRARY       : DERIVATIVES
'GROUP         : ASIAN
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function BINOMIAL_ASIAN_OPTION_PRICE_FUNC(ByVal SPOT As Double, _
ByVal STRIKE As Double, _
ByVal EXPIRATION As Double, _
ByVal RISK_FREE_RATE As Double, _
ByVal VOLATILITY As Double, _
Optional ByVal CONSTANT As Double = 0.1, _
Optional ByVal nLOOPS As Long = 60, _
Optional ByVal EXERCISE_TYPE As Integer = 0, _
Optional ByVal OPTION_FLAG As Integer = 1) 'As Double

'SPOT - spot price (50)
'STRIKE - option strike (50)
'EXPIRATION - option maturity (1)
'RISK_FREE_RATE - risk free rate (0.1)
'VOLATILITY - volatility (0.3)
'CONSTANT - constant value for discreting possible choices of f (0.1)
'nLOOPS - no of time steps for the binomial tree (60)
'EXERCISE_TYPE - use 0 for American else for european (0)
'OPTION_FLAG - use 1 for call and -1 for put option (1)

'Option Price    5.359642991


Dim i As Long
Dim j As Long
Dim k As Long

Dim ii As Long
Dim jj As Long
Dim kk As Long

Dim iii As Long 'Min
Dim jjj As Long 'Max

Dim A_VAL As Double
Dim D_VAL As Double
Dim F_VAL As Double
Dim U_VAL As Double

Dim X_VAL As Double
Dim Y_VAL As Double

Dim DT_VAL As Double
Dim PU_VAL As Double
Dim PD_VAL As Double

Dim SU_VAL As Double
Dim FU_VAL As Double

Dim VD_VAL As Double
Dim VU_VAL As Double

Dim SD_VAL As Double
Dim FD_VAL As Double

Dim IM_VAL As Double
Dim PV_VAL As Double
Dim NEW_VAL As Double
Dim TEMP_VAL As Double

Dim AVG_MIN_VAL As Double
Dim AVG_MAX_VAL As Double

Dim SPOT_MIN_VAL As Double
Dim SPOT_MAX_VAL As Double

Dim PREV_AVG_MIN_VAL As Double
Dim PREV_AVG_MAX_VAL As Double

Dim FTREE_ARR As Variant 'stores vector of possible value of running average till current node
Dim FVEC_ARR As Variant
Dim PRICE_ARR As Variant 'stores vector of stock prices at a time
Dim TREE_ARR As Variant 'stores stock prices for entire tree

Dim AVG_ARR As Variant 'running average values to consider for calculating set of option prices
Dim TIME_ARR As Variant 'stores option values at all nodes of the tree
Dim NODE_ARR As Variant 'stores list of vectors of option prices at a given time

Dim PREV_ARR As Variant 'PrevAverageVec
Dim NEXT_ARR As Variant 'VTimeVecNext

Dim TEMP_ARR As Variant 'Svecnext
Dim XTEMP_ARR As Variant
Dim YTEMP_ARR As Variant

On Error GoTo ERROR_LABEL

If OPTION_FLAG <> 1 Then: OPTION_FLAG = -1

DT_VAL = EXPIRATION / nLOOPS
U_VAL = Exp(VOLATILITY * DT_VAL ^ 0.5) 'size of up jump
D_VAL = Exp(-VOLATILITY * DT_VAL ^ 0.5) 'size of down jump
A_VAL = Exp(RISK_FREE_RATE * DT_VAL)
PU_VAL = (A_VAL - D_VAL) / (U_VAL - D_VAL) 'probability of up jump
PD_VAL = 1 - PU_VAL 'probability of down jump

'Step 1 : calculate stock prices

ReDim TREE_ARR(0 To nLOOPS) 'stores stock prices for entire tree
For i = 0 To nLOOPS
    ReDim PRICE_ARR(0 To i) 'stores vector of stock prices at a time
    For j = 0 To i
        PRICE_ARR(j) = SPOT * U_VAL ^ j * D_VAL ^ (i - j)
    Next j
    TREE_ARR(i) = PRICE_ARR
Next i

'Step 2 : calculate values of F_VAL at each node
ReDim FTREE_ARR(0 To nLOOPS) 'stores vector of possible value of running average till current node
ReDim FVEC_ARR(0 To 0)

FVEC_ARR(0) = SPOT
FTREE_ARR(0) = FVEC_ARR

For i = 1 To nLOOPS
'find minimum and maximum values of running average
'at each time step
    
    PRICE_ARR = TREE_ARR(i)
    SPOT_MAX_VAL = PRICE_ARR(UBound(PRICE_ARR, 1))
    SPOT_MIN_VAL = PRICE_ARR(0)
    PREV_ARR = FTREE_ARR(i - 1)
    PREV_AVG_MAX_VAL = PREV_ARR(UBound(PREV_ARR, 1))
    PREV_AVG_MIN_VAL = PREV_ARR(0)
    AVG_MAX_VAL = (PREV_AVG_MAX_VAL * i + SPOT_MAX_VAL) / (i + 1)
    AVG_MIN_VAL = (PREV_AVG_MIN_VAL * i + SPOT_MIN_VAL) / (i + 1)
    'now find integer values of m which cover min and max average values
    
    TEMP_VAL = Log(AVG_MAX_VAL / SPOT) / CONSTANT
    jjj = Int(Abs(TEMP_VAL)) + 1 'WorksheetFunction.Floor(TEMP_VAL, 1) + 1
    TEMP_VAL = Log(AVG_MIN_VAL / SPOT) / CONSTANT
    iii = Int(Abs(TEMP_VAL)) + 1 'WorksheetFunction.Floor(Abs(TEMP_VAL), 1) + 1
    kk = jjj + iii + 1
    
    ReDim FVEC_ARR(0 To kk - 1)
    k = -iii
    For j = 0 To kk - 1
        FVEC_ARR(j) = SPOT * Exp(k * CONSTANT)
        k = k + 1
    Next j
    FTREE_ARR(i) = FVEC_ARR
Next i

'Step 3 : Do backward recursion of the tree
'initialize option values at maturity
FVEC_ARR = FTREE_ARR(nLOOPS)
'running average values to consider for calculating set of option prices
ReDim AVG_ARR(0 To nLOOPS)
'stores option values at all nodes of the tree
ReDim TIME_ARR(0 To nLOOPS)
'stores list of vectors of option prices at a given time
ReDim NODE_ARR(0 To UBound(FVEC_ARR, 1))

For j = 0 To UBound(FVEC_ARR, 1) 'loop over average values
    NODE_ARR(j) = MAXIMUM_FUNC(FVEC_ARR(j) - STRIKE, 0)
Next j
For i = 0 To nLOOPS 'loop over nodes at a given time
    TIME_ARR(i) = NODE_ARR
Next i
AVG_ARR(nLOOPS) = TIME_ARR

For i = nLOOPS - 1 To 0 Step -1
    TEMP_ARR = TREE_ARR(i + 1)
    
    ReDim TIME_ARR(0 To i)
    NEXT_ARR = AVG_ARR(i + 1)
    FVEC_ARR = FTREE_ARR(i)
    XTEMP_ARR = FTREE_ARR(i + 1)
    
    For j = 0 To i
        ReDim NODE_ARR(0 To UBound(FVEC_ARR, 1))
        For jj = 0 To UBound(FVEC_ARR, 1)
            'calculate option price using F_VAL at current node and SU_VAL
            F_VAL = FVEC_ARR(jj) 'running average
            'find running average at next time node of up-jump
            SU_VAL = TEMP_ARR(j + 1)
            FU_VAL = (F_VAL * (i + 1) + SU_VAL) / (i + 2)
            YTEMP_ARR = NEXT_ARR(j + 1) 'vector of option prices
            'get option value to the next timestep
            X_VAL = CDbl(FU_VAL)
            Y_VAL = 0
            GoSub 1983
            VU_VAL = Y_VAL
            
            'find running average at next time node of down-jump
            SD_VAL = TEMP_ARR(j)
            FD_VAL = (F_VAL * (i + 1) + SD_VAL) / (i + 2)
            YTEMP_ARR = NEXT_ARR(j) 'vector of option prices
            'get option value to the next timestep
            X_VAL = CDbl(FD_VAL)
            Y_VAL = 0
            GoSub 1983
            VD_VAL = Y_VAL
            
            'NEW_VAL = Exp(-RISK_FREE_RATE * DT_VAL) * (VU_VAL * PU_VAL + VD_VAL * PD_VAL)
            PV_VAL = Exp(-RISK_FREE_RATE * DT_VAL) * (VU_VAL * PU_VAL + VD_VAL * PD_VAL)
            IM_VAL = OPTION_FLAG * (F_VAL - STRIKE)
            If EXERCISE_TYPE = 0 Then 'american
                NEW_VAL = MAXIMUM_FUNC(PV_VAL, IM_VAL)
            Else 'european
                NEW_VAL = MAXIMUM_FUNC(PV_VAL, 0)
            End If
            NODE_ARR(jj) = NEW_VAL
        Next jj
        TIME_ARR(j) = NODE_ARR
    Next j
    AVG_ARR(i) = TIME_ARR
Next i

TIME_ARR = AVG_ARR(0)
NODE_ARR = TIME_ARR(0)
BINOMIAL_ASIAN_OPTION_PRICE_FUNC = NODE_ARR(0)
'---------------------------------------------------------------------------------
Exit Function
'---------------------------------------------------------------------------------
1983: 'Returns an interpolated value of X_VAL doing a lookup of
'XTEMP_ARR
'---------------------------------------------------------------------------------
If ((X_VAL < XTEMP_ARR(LBound(XTEMP_ARR))) Or _
    (X_VAL > XTEMP_ARR(UBound(XTEMP_ARR)))) Then: GoTo ERROR_LABEL
' X_VAL is out of bound"

If XTEMP_ARR(LBound(XTEMP_ARR)) = X_VAL Then
    Y_VAL = YTEMP_ARR(LBound(YTEMP_ARR))
    Return
End If
For ii = LBound(XTEMP_ARR) To UBound(XTEMP_ARR)
    If XTEMP_ARR(ii) >= X_VAL Then
        Y_VAL = YTEMP_ARR(ii - 1) + (X_VAL - XTEMP_ARR(ii - 1)) / (XTEMP_ARR(ii) - _
                XTEMP_ARR(ii - 1)) * (YTEMP_ARR(ii) - YTEMP_ARR(ii - 1))
        Return
    End If
Next ii
'---------------------------------------------------------------------------------
Exit Function
'---------------------------------------------------------------------------------
ERROR_LABEL:
BINOMIAL_ASIAN_OPTION_PRICE_FUNC = Err.number
End Function
