Attribute VB_Name = "FINAN_DERIV_BINOM_CLIQ_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : EUROPEAN_CLIQUET_OPTION_FUNC
'DESCRIPTION   : Binomial model for pricing European reset options with
'risk-neutral probabilities. In this function instead of setting
'an array for the Backward Valuation Binomial, i used the Factorial
'function to derive the closed-form solution.

'LIBRARY       : DERIVATIVES
'GROUP         : CLIQUET
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function EUROPEAN_CLIQUET_OPTION_FUNC(ByVal SPOT As Double, _
ByVal STRIKE As Double, _
ByVal EXPIRATION As Double, _
ByVal RESET As Long, _
ByVal RATE As Double, _
ByVal DIVD As Double, _
ByVal SIGMA As Double, _
ByVal STEPS As Long, _
Optional ByVal OPTION_FLAG As Integer = 1)

'RESET: Years to reset

Dim i As Long
Dim j As Long

Dim SPOT_RESET As Double 'calculating stock prices at the reset date
Dim SPOT_EXPIRATION As Double 'calculating stock prices at the MATURITY date

Dim TEMP_FACT As Double 'calculating PAYOFF probabilities
Dim OPTION_VAL As Double 'calculating expected PAYOFFs

Dim DISCOUNT_FACTOR As Double
  
Dim DELTA_TIME As Double  'time step in the tree

Dim GROWTH_FACTOR As Double 'the cost of carry term
Dim UP_STEP_SIZE As Double
Dim DOWN_STEP_SIZE As Double

Dim PROB_UP_MOVE As Double 'risk-neutral UP probability
Dim PROB_DOWN_MOVE As Double

On Error GoTo ERROR_LABEL

If OPTION_FLAG <> 1 Then OPTION_FLAG = -1

DELTA_TIME = EXPIRATION / STEPS
DISCOUNT_FACTOR = Exp(-RATE * DELTA_TIME)
RESET = Int(RESET / DELTA_TIME)

'the default is a call option, unless a "PUT" is given to indicate put option
    
' the stock price S move to S*UP_STEP_SIZE and S*DOWN_STEP_SIZE in the next _
  time step of the tree

' UP_STEP_SIZE and DOWN_STEP_SIZE and probability of going up or down are _
  calculated in the following
' This is the standard Cox - Rubinstein tree setup

UP_STEP_SIZE = Exp(SIGMA * DELTA_TIME ^ (0.5))
DOWN_STEP_SIZE = 1 / UP_STEP_SIZE
        
GROWTH_FACTOR = Exp((RATE - DIVD) * DELTA_TIME) 'cost of carry term
PROB_UP_MOVE = (GROWTH_FACTOR - DOWN_STEP_SIZE) / (UP_STEP_SIZE - DOWN_STEP_SIZE)
PROB_DOWN_MOVE = 1 - PROB_UP_MOVE
    
For j = 0 To RESET
    
    SPOT_RESET = SPOT * UP_STEP_SIZE ^ j * DOWN_STEP_SIZE ^ (RESET - j)
    
    For i = j To (STEPS - RESET + j)
        
        SPOT_EXPIRATION = SPOT * UP_STEP_SIZE ^ i * DOWN_STEP_SIZE ^ (STEPS - i)
        
        TEMP_FACT = (FACTORIAL_FUNC(RESET) * FACTORIAL_FUNC(STEPS - RESET)) / _
        (FACTORIAL_FUNC(j) * FACTORIAL_FUNC(RESET - j) * _
        FACTORIAL_FUNC(i - j) * FACTORIAL_FUNC(STEPS - RESET - i + j)) * _
        PROB_UP_MOVE ^ i * (1 - PROB_UP_MOVE) ^ (STEPS - i)
        
        'The use of the Boolean variable OPTION_FLAG (which is set to 1 for call
        'and -1 for put) helps prevent unnecessary "IF... THEN..." checks.
        
        OPTION_VAL = OPTION_VAL + TEMP_FACT * _
        MAXIMUM_FUNC(MAXIMUM_FUNC(OPTION_FLAG * (SPOT_EXPIRATION - STRIKE), _
        OPTION_FLAG * (SPOT_EXPIRATION - SPOT_RESET)), 0)
        
    Next i
Next j

EUROPEAN_CLIQUET_OPTION_FUNC = Exp(-RATE * EXPIRATION) * OPTION_VAL 'Option Value

Exit Function
ERROR_LABEL:
EUROPEAN_CLIQUET_OPTION_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : EUROPEAN_CLIQUET_OPTION_TREE_FUNC
'DESCRIPTION   : Function for pricing European reset options with risk-neutral
'probabilities
'LIBRARY       : DERIVATIVES
'GROUP         : CLIQUET
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function EUROPEAN_CLIQUET_OPTION_TREE_FUNC(ByVal SPOT As Double, _
ByVal STRIKE As Double, _
ByVal EXPIRATION As Double, _
ByVal RESET As Long, _
ByVal RATE As Double, _
ByVal DIVD As Double, _
ByVal SIGMA As Double, _
ByVal STEPS As Long, _
Optional ByVal OPTION_FLAG As Integer = 1, _
Optional ByVal EXERCISE_TYPE As Integer = 0)
                            
Dim i As Long
Dim j As Long
Dim k As Long 'number of steps from reset to MATURITY

Dim GROWTH_FACTOR As Double
Dim UP_STEP_SIZE As Double
Dim DOWN_STEP_SIZE As Double
Dim PROB_UP_MOVE As Double
Dim DELTA_TIME As Double

Dim DISCOUNT_FACTOR As Double
Dim RESET_EXPIRATION As Double 'time from reset to MATURITY

Dim SPOT_VECTOR As Variant 'stock prices at reset date
Dim STRIKE_VECTOR As Variant 'new strike after reset
Dim OPTION_VECTOR As Variant 'at reset time vanillas values with STEPS-RESET steps to MATURITY

On Error GoTo ERROR_LABEL

'--------------------------Two-year Reset Put on NICO Index----------------------
'   Option Buyer XXX
'   Option Seller YYY
'   Start Date dd/mm/yyyy
'   Maturity Date Start date + 2 years
'   Option Seller Pays at Maturity:
'       St – ST if St > k, ST = St
'       k – ST if St = k, ST = k
'
'       0 if (St > k and ST > St) or (St = k and ST > k):
'       where, k is the original exercise price (which is the closing stock or
'       index price on the date of issuance), St is the closing stock or index
'       price on the reset date, and ST is the closing stock/index price on the
'       EXPIRATION date.
'--------------------------------------------------------------------------------

STEPS = Fix(STEPS / 2#) * 2# 'AVOID ODD ENTRIES
RESET = STEPS * RESET

ReDim OPTION_VECTOR(0 To STEPS, 1 To 1)
ReDim SPOT_VECTOR(0 To RESET, 1 To 1)
ReDim STRIKE_VECTOR(0 To RESET, 1 To 1)
                           
DELTA_TIME = EXPIRATION / STEPS
UP_STEP_SIZE = Exp(SIGMA * DELTA_TIME ^ (0.5))
DOWN_STEP_SIZE = 1 / UP_STEP_SIZE

RESET_EXPIRATION = (STEPS - RESET) * EXPIRATION / STEPS 'time from reset to maturity
k = STEPS - RESET 'number of steps from reset to maturity

'Option valuation at reset date at each stock price node
For i = 0 To RESET
    SPOT_VECTOR(i, 1) = SPOT * UP_STEP_SIZE ^ i * _
    DOWN_STEP_SIZE ^ (RESET - i) 'stock prices at reset date
    STRIKE_VECTOR(i, 1) = OPTION_FLAG * MINIMUM_FUNC(OPTION_FLAG * SPOT_VECTOR(i, 1), _
    OPTION_FLAG * STRIKE) 'new strike after reset
    
    OPTION_VECTOR(i, 1) = BINOMIAL_TREE_CONSTANT_PRICE_FUNC(SPOT_VECTOR(i, 1), STRIKE_VECTOR(i, 1), _
    RESET_EXPIRATION, RATE, DIVD, SIGMA, k, OPTION_FLAG, EXERCISE_TYPE, 0, 0)(0, 1)
    'at reset time vanillas values with k to maturity
Next i

GROWTH_FACTOR = Exp((RATE - DIVD) * DELTA_TIME) 'cost of carry term
PROB_UP_MOVE = (GROWTH_FACTOR - DOWN_STEP_SIZE) / (UP_STEP_SIZE - _
DOWN_STEP_SIZE) 'risk-neutral UP probability
DISCOUNT_FACTOR = Exp((-RATE) * DELTA_TIME)


Select Case EXERCISE_TYPE 'Backwards option valuation
Case 0 ', "EURO"
    For j = RESET - 1 To 0 Step -1
        For i = 0 To j
             OPTION_VECTOR(i, 1) = (PROB_UP_MOVE * OPTION_VECTOR(i + 1, 1) + _
             (1 - PROB_UP_MOVE) * _
             OPTION_VECTOR(i, 1)) * DISCOUNT_FACTOR
        Next i
    Next j
Case Else '1, "AMER"
    For j = RESET - 1 To 0 Step -1
        For i = 0 To j
             OPTION_VECTOR(i, 1) = MAXIMUM_FUNC((OPTION_FLAG * (SPOT * UP_STEP_SIZE ^ i * _
             DOWN_STEP_SIZE ^ (j - i) - STRIKE)), (PROB_UP_MOVE * OPTION_VECTOR(i + 1, 1) + _
             (1 - PROB_UP_MOVE) _
             * OPTION_VECTOR(i, 1)) * DISCOUNT_FACTOR)
        Next i
    Next j
End Select

EUROPEAN_CLIQUET_OPTION_TREE_FUNC = OPTION_VECTOR

Exit Function
ERROR_LABEL:
EUROPEAN_CLIQUET_OPTION_TREE_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : DELTA_CLIQUET_FUNCTION
'DESCRIPTION   : DELTA_CLIQUET_FUNCTION
'LIBRARY       : DERIVATIVES
'GROUP         : CLIQUET
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function EUROPEAN_CLIQUET_OPTION_DELTA_FUNC(ByVal DELTA_SPOT As Double, _
ByVal SPOT As Double, _
ByVal STRIKE As Double, _
ByVal EXPIRATION As Double, _
ByVal RESET As Long, _
ByVal RATE As Double, _
ByVal DIVD As Double, _
ByVal SIGMA As Double, _
ByVal STEPS As Long, _
Optional ByVal OPTION_FLAG As Integer = 1, _
Optional ByVal EXERCISE_TYPE As Integer = 0)

'DELTA_SPOT = 0.0001

Dim FIRST_VAL As Double
Dim SECOND_VAL As Double

On Error GoTo ERROR_LABEL

FIRST_VAL = EUROPEAN_CLIQUET_OPTION_TREE_FUNC(SPOT + DELTA_SPOT, STRIKE, _
EXPIRATION, RESET, RATE, DIVD, SIGMA, STEPS, OPTION_FLAG, EXERCISE_TYPE)(0, 1)

SECOND_VAL = EUROPEAN_CLIQUET_OPTION_TREE_FUNC(SPOT, STRIKE, _
EXPIRATION, RESET, RATE, DIVD, SIGMA, STEPS, OPTION_FLAG, EXERCISE_TYPE)(0, 1)

EUROPEAN_CLIQUET_OPTION_DELTA_FUNC = (FIRST_VAL - SECOND_VAL) / DELTA_SPOT
                                    
Exit Function
ERROR_LABEL:
EUROPEAN_CLIQUET_OPTION_DELTA_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : EUROPEAN_CLIQUET_OPTION_SIMULATION_FUNC

'DESCRIPTION   : Monte Carlo Simulation for European Cliquet Options Valuation
'Cliquet Options Monte-Carlo Simulation Model
'this model prices only European cliquets and is not as
'accurate as our backward valuation method.

'The greatest advantage of Monte-Carlo simulation pricing is that
'this model can be most easily adjusted for almost any other type
'of cliquet options (for instance: multiple-reset compound ratchets,
'coupe options, barrier cliquets, etc.) as well as for many other
'different path-dependent exotics.

'The greatest advantage of using MC simulation, as opposed to closed-form
'solutions and other numerical and analytical approaches, is its tremendous
'flexibility. It is possible to adjust Monte-Carlo simulation to price almost
'any kind of exotic. The different cliquet option types mentioned above
'(including multiple reset dates) could be priced using MC
'simulation.

'However, there are many tradeoffs involved in using MC simulations. Among the
'disadvantages of Monte-Carlo:
'   The inability to value American options (though there're several approaches
'   that deal with this problem).

'Relatively poor accuracy on one hand, and an extremely large number of calculations
'needed to reach a good accuracy, on the other.

'LIBRARY       : DERIVATIVES
'GROUP         : CLIQUET
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************


Function EUROPEAN_CLIQUET_OPTION_SIMULATION_FUNC(ByVal SPOT As Double, _
ByVal STRIKE As Double, _
ByVal EXPIRATION As Double, _
ByVal RESET As Long, _
ByVal RATE As Double, _
ByVal DIVD As Double, _
ByVal SIGMA As Double, _
ByVal STEPS As Long, _
ByVal NTRIALS As Long, _
ByVal nLOOPS As Long, _
Optional ByVal OPTION_FLAG As Integer = 1, _
Optional ByVal RANDOM_FLAG As Boolean = True)

Dim i As Long
Dim j As Long
Dim k As Long

Dim TEMP_RND As Double
Dim TEMP_SUM As Double
Dim DELTA_TIME As Double

Dim OPTION_VAL As Double
Dim STRIKE_VAL As Double

Dim UP_STEP_SIZE As Double
Dim DOWN_STEP_SIZE As Double
Dim PROB_UP_MOVE As Double
Dim GROWTH_FACTOR As Double

Dim TEMP_VECTOR As Variant
Dim OPTION_VECTOR As Variant

On Error GoTo ERROR_LABEL

If RANDOM_FLAG = True Then: Randomize

ReDim OPTION_VECTOR(0 To STEPS, 1 To 1)
ReDim TEMP_VECTOR(1 To nLOOPS, 1 To 1)

RESET = RESET * STEPS
DELTA_TIME = EXPIRATION / STEPS
UP_STEP_SIZE = Exp(SIGMA * Sqr(DELTA_TIME))
DOWN_STEP_SIZE = 1 / UP_STEP_SIZE
GROWTH_FACTOR = Exp((RATE - DIVD) * DELTA_TIME)
PROB_UP_MOVE = (GROWTH_FACTOR - DOWN_STEP_SIZE) / (UP_STEP_SIZE - DOWN_STEP_SIZE)

For k = 1 To nLOOPS
    
    TEMP_SUM = 0
    OPTION_VAL = 0
    STRIKE_VAL = 0
    
    For j = 1 To NTRIALS 'OPTION_VECTOR path generation

'At each time step a number between 0 and 1 is drawn at random from a uniform
'distribution. If this number is less than p, the stock takes an up step, if
'the random number is larger than p, the stock goes down one step. On the
'reset day m, the new strike is determined and the stock price continues its
'random walk. On the maturity date n, the payoff is determined and discounted
'by the risk-free rate to get its present value. The procedure repeats itself
'a number of times (5,000 works good enough). Then, a simple arithmetic average
'of all discounted payoffs is calculated, and this is a European cliquet
'value under Monte-Carlo valuation.
        
        For i = 1 To STEPS
            OPTION_VECTOR(0, 1) = SPOT 'UP/Down determining
            TEMP_RND = Rnd
            If TEMP_RND < PROB_UP_MOVE Then
                OPTION_VECTOR(i, 1) = OPTION_VECTOR(i - 1, 1) * UP_STEP_SIZE
            Else
                OPTION_VECTOR(i, 1) = OPTION_VECTOR(i - 1, 1) * DOWN_STEP_SIZE
            End If
        Next i
'-------------------------------------Strike reset
        STRIKE_VAL = OPTION_FLAG * MINIMUM_FUNC(OPTION_FLAG * OPTION_VECTOR(RESET, 1), _
        OPTION_FLAG * STRIKE)
'--------------------------option value in one simulation
        OPTION_VAL = MAXIMUM_FUNC(0, OPTION_FLAG * (OPTION_VECTOR(STEPS, 1) - _
                    STRIKE_VAL))
'-------------------sum of option values in all simulations
        TEMP_SUM = TEMP_SUM + OPTION_VAL
    Next j
    TEMP_VECTOR(k, 1) = (TEMP_SUM * Exp(-RATE * EXPIRATION)) / NTRIALS
    'average cliquets discounted value
Next k

EUROPEAN_CLIQUET_OPTION_SIMULATION_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
EUROPEAN_CLIQUET_OPTION_SIMULATION_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : BINOMIAL_TREE_CONSTANT_PRICE_FUNC

'DESCRIPTION   : The following model is built under the assumption that cash paid to
'owners of the underlying, such as dividends and interest, are paid
'continuously at constant risk-free over the life of the option. This
'assumption is relatively accurate for valuing puts generally, and calls
'on bonds, commodities, currencies and stock index portfolios.

'LIBRARY       : DERIVATIVES
'GROUP         : CLIQUET
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Private Function BINOMIAL_TREE_CONSTANT_PRICE_FUNC(ByVal S_VAL As Double, _
ByVal K_VAL As Double, _
ByVal T_VAL As Double, _
ByVal RF_VAL As Double, _
ByVal DY_VAL As Double, _
ByVal V_VAL As Double, _
ByVal N_VAL As Long, _
Optional ByVal OT_VAL As Integer = 1, _
Optional ByVal ET_VAL As Integer = 0, _
Optional ByVal MODEL As Integer = 0, _
Optional ByVal OUTPUT As Integer = 0)

'S_VAL: Underlying asset current price.

'K_VAL: Price at which the underlying can be bought (for calls) or
'sold (for puts).

'T_VAL: Maturity scaled on an annualized basis

'RF_VAL: The cost of funds for the number of days to MATURITY
'It should corresponds to the period to the record date, not the option MATURITY

'DY_VAL: Yield paid on the underlying asset matching
'the option's MATURITY.

'V_VAL: Annualized V_VAL of the underlying asset price

'N_VAL --> Number of N_VAL (e.g. One-month intervals)

'OT_VAL: 1 for call, and -1 for put.
'ET_VAL: 0 For European, 1 For American, 2 For Early Exercise Condition
'MODEL: 0 for Standard Cox - Rubinstein tree setup, and 1 for Binomial JR
    
Dim i As Long
Dim j As Long
    
Dim DF_VAL As Double 'Discount Factor
Dim DT_VAL As Double  'time step in the tree
    
Dim GF_VAL As Double 'Growth Factor
Dim US_VAL As Double 'Up Step
Dim DS_VAL As Double 'Down Step
Dim UP_VAL As Double 'Up Prob
Dim DN_VAL As Double 'Down prob
    
Dim NODES1_ARR As Variant 'Stock Values
Dim NODES2_ARR As Variant 'Option Values
Dim NODES3_ARR As Variant 'Early Vector
    
On Error GoTo ERROR_LABEL
    
If OT_VAL <> 1 Then: OT_VAL = -1 'Put

DT_VAL = T_VAL / N_VAL
DF_VAL = Exp(-RF_VAL * DT_VAL)
        
' the stock price S move to S*US_VAL and S*DS_VAL in the
' next time step of the tree
    
'-------------------------------------------------------------------------------------------------------------------------
Select Case MODEL
'-------------------------------------------------------------------------------------------------------------------------
Case 0 'Standard Cox - Rubinstein tree setup
'-------------------------------------------------------------------------------------------------------------------------
    US_VAL = Exp(V_VAL * DT_VAL ^ (0.5))
    DS_VAL = 1 / US_VAL
    GF_VAL = Exp((RF_VAL - DY_VAL) * DT_VAL) 'cost of carry term
    UP_VAL = (GF_VAL - DS_VAL) / (US_VAL - DS_VAL)
    DN_VAL = 1 - UP_VAL
'-------------------------------------------------------------------------------------------------------------------------
Case Else 'JR
'-------------------------------------------------------------------------------------------------------------------------
    US_VAL = Exp(((RF_VAL - DY_VAL) - 0.5 * V_VAL ^ 2) * DT_VAL + V_VAL * DT_VAL ^ (0.5))
    DS_VAL = Exp(((RF_VAL - DY_VAL) - 0.5 * V_VAL ^ 2) * DT_VAL - V_VAL * DT_VAL ^ (0.5))
    UP_VAL = 0.5
    DN_VAL = (1 - UP_VAL)
'-------------------------------------------------------------------------------------------------------------------------
End Select
'-------------------------------------------------------------------------------------------------------------------------

ReDim NODES1_ARR(0 To N_VAL): ReDim NODES2_ARR(0 To N_VAL): ReDim NODES3_ARR(0 To N_VAL)
' Initialise asset prices at MATURITY
NODES1_ARR(0) = S_VAL * (DS_VAL ^ N_VAL)
For i = 1 To N_VAL
    NODES1_ARR(i) = NODES1_ARR(i - 1) * US_VAL / DS_VAL
Next i

' initialise option values at MATURITY
For i = 0 To N_VAL
    NODES2_ARR(i) = MAXIMUM_FUNC(0#, OT_VAL * NODES1_ARR(i) - OT_VAL * K_VAL)
    NODES3_ARR(i) = 0
Next i
    
'-------------------------------------------------------------------------------------------------------------------------
Select Case ET_VAL     ' stepping back the tree
'-------------------------------------------------------------------------------------------------------------------------
Case 0 ', "EURO"
'-------------------------------------------------------------------------------------------------------------------------
    For j = N_VAL - 1 To 0 Step -1
        For i = 0 To j
            NODES2_ARR(i) = DF_VAL * (UP_VAL * NODES2_ARR(i + 1) + DN_VAL * NODES2_ARR(i))
        Next i
    Next j
'-------------------------------------------------------------------------------------------------------------------------
Case 1 ', "AMER"
'-------------------------------------------------------------------------------------------------------------------------
    For j = N_VAL - 1 To 0 Step -1
        For i = 0 To j
            NODES2_ARR(i) = DF_VAL * (UP_VAL * _
            NODES2_ARR(i + 1) + DN_VAL * NODES2_ARR(i))
            NODES1_ARR(i) = NODES1_ARR(i) * US_VAL
            NODES2_ARR(i) = MAXIMUM_FUNC(NODES2_ARR(i), OT_VAL * NODES1_ARR(i) - OT_VAL * K_VAL)
        Next i
    Next j
'-------------------------------------------------------------------------------------------------------------------------
Case Else 'CALCULATING EARLY EXERCISE CONDITIONS
'-------------------------------------------------------------------------------------------------------------------------
    For j = N_VAL - 1 To 0 Step -1
        For i = 0 To j
            NODES2_ARR(i) = DF_VAL * (UP_VAL * NODES2_ARR(i + 1) + DN_VAL * NODES2_ARR(i))
            NODES1_ARR(i) = NODES1_ARR(i) * US_VAL
            If NODES2_ARR(i) > (OT_VAL * NODES1_ARR(i) - OT_VAL * K_VAL) Then
                NODES3_ARR(i) = (UP_VAL * NODES3_ARR(i + 1) + DN_VAL * NODES3_ARR(i)) + DT_VAL
            Else
                NODES3_ARR(i) = 0
            End If
        Next i
    Next j
'-------------------------------------------------------------------------------------------------------------------------
End Select
'-------------------------------------------------------------------------------------------------------------------------
    
'-------------------------------------------------------------------------------------------------------------------------
Select Case OUTPUT
'-------------------------------------------------------------------------------------------------------------------------
Case 0
'-------------------------------------------------------------------------------------------------------------------------
    BINOMIAL_TREE_CONSTANT_PRICE_FUNC = NODES2_ARR
'-------------------------------------------------------------------------------------------------------------------------
Case 1
'-------------------------------------------------------------------------------------------------------------------------
    If ET_VAL = 2 Then
        BINOMIAL_TREE_CONSTANT_PRICE_FUNC = NODES3_ARR
    Else
        BINOMIAL_TREE_CONSTANT_PRICE_FUNC = "PLEASE CHOOSE ET_VAL = 2"
    End If
'-------------------------------------------------------------------------------------------------------------------------
Case Else
'-------------------------------------------------------------------------------------------------------------------------
    BINOMIAL_TREE_CONSTANT_PRICE_FUNC = NODES1_ARR
'-------------------------------------------------------------------------------------------------------------------------
End Select
'-------------------------------------------------------------------------------------------------------------------------
    
Exit Function
ERROR_LABEL:
BINOMIAL_TREE_CONSTANT_PRICE_FUNC = Err.number
End Function

'------------------------------------------------------------------------------------
'SOURCE: http://www.global-derivatives.com/download/files.php?func=download&id=114
'------------------------------------------------------------------------------------
'Recent turmoil in financial equity markets has caused an increase in demand for
'products that reduce downside risk while still offering upside potential. Reset
'options, also termed cliquet, ratchet, and strike reset options, provide a
'product structured to meet that demand. Reset puts, appeal to large pension funds,
'portfolio insurers as well as retail investors.

'The first cliquet options to be traded on a public exchange were S&P 500 bear market
'warrant with a periodic reset. These started trading on the Chicago Board of Options
'Exchange in 1996. These reset warrants on the S&P 500 index work like regular equity
'or index puts except that the exercise price is reset at a higher level if the index
'level is above the original exercise price on the reset date.
'------------------------------------------------------------------------------------
'Reset Options: Reset option can change the strike at some specified times when
'the option is out-of-the-money. It is common to change to strike to the prevailing
'spot price so that option is "enhanced" to be at-the-money. Reset features have
'been used in equity derivatives such as warrants and convertibles (especially in
'Japan). It is attractive to the investors while the extra (optionality) cost is
'not too high. The hedging of such options is difficult since the delta has some
'discontinuities like the barrier options.
'------------------------------------------------------------------------------------

