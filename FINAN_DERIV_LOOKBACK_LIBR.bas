Attribute VB_Name = "FINAN_DERIV_LOOKBACK_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : FLOATING_STRIKE_LOOKBACK_OPTION_FUNC
'DESCRIPTION   : Floating strike lookback options
'LIBRARY       : DERIVATIVES
'GROUP         : LOOKBACK
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function FLOATING_STRIKE_LOOKBACK_OPTION_FUNC(ByVal SPOT As Double, _
ByVal MIN_SPOT As Double, _
ByVal MAX_SPOT As Double, _
ByVal EXPIRATION As Double, _
ByVal RATE As Double, _
ByVal CARRY_COST As Double, _
ByVal SIGMA As Double, _
Optional ByVal OPTION_FLAG As Integer = 1, _
Optional ByVal CND_TYPE As Integer = 0)

'------------------------------------------------------
'MIN_SPOT: Observed minimum
'MAX_SPOT: Observed Maximum
'EXPIRATION: Time to maturity
'------------------------------------------------------

Dim D1_VAL As Double
Dim D2_VAL As Double
Dim TEMP_SPOT As Double
    
On Error GoTo ERROR_LABEL

If OPTION_FLAG <> 1 Then: OPTION_FLAG = -1
    
Select Case OPTION_FLAG
Case 1 ', "c", "call" '---> Call on the minimum
    TEMP_SPOT = MIN_SPOT
Case Else '-1', "p", "put" '---> Put on the maximum
    TEMP_SPOT = MAX_SPOT
End Select
    
D1_VAL = (Log(SPOT / TEMP_SPOT) + (CARRY_COST + SIGMA ^ 2 / 2) * EXPIRATION) / _
(SIGMA * Sqr(EXPIRATION))
     
D2_VAL = D1_VAL - SIGMA * Sqr(EXPIRATION)

Select Case OPTION_FLAG
'----------------------------------------------------------------------------------
Case 1 ', "c", "call" '---> Call on the minimum
'----------------------------------------------------------------------------------
        FLOATING_STRIKE_LOOKBACK_OPTION_FUNC = SPOT * Exp((CARRY_COST - RATE) * EXPIRATION) * _
            CND_FUNC(D1_VAL, CND_TYPE) - TEMP_SPOT * Exp(-RATE * EXPIRATION) * _
            CND_FUNC(D2_VAL, CND_TYPE) + Exp(-RATE * EXPIRATION) * SIGMA ^ 2 / _
            (2 * CARRY_COST) * SPOT * _
            ((SPOT / TEMP_SPOT) ^ (-2 * CARRY_COST / SIGMA ^ 2) * CND_FUNC(-D1_VAL + 2 * _
            CARRY_COST / SIGMA * Sqr(EXPIRATION), CND_TYPE) - _
            Exp(CARRY_COST * EXPIRATION) _
            * CND_FUNC(-D1_VAL, CND_TYPE))
'----------------------------------------------------------------------------------
Case Else '-1 ', "p", "put" '---> Put on the maximum
'----------------------------------------------------------------------------------
        FLOATING_STRIKE_LOOKBACK_OPTION_FUNC = TEMP_SPOT * Exp(-RATE * EXPIRATION) * _
            CND_FUNC(-D2_VAL, CND_TYPE) - SPOT * Exp((CARRY_COST - RATE) * _
            EXPIRATION) * CND_FUNC(-D1_VAL, CND_TYPE) + _
            Exp(-RATE * EXPIRATION) * SIGMA ^ 2 / (2 * CARRY_COST) * SPOT * _
            (-(SPOT / TEMP_SPOT) ^ (-2 * CARRY_COST / SIGMA ^ 2) * _
            CND_FUNC(D1_VAL - 2 * CARRY_COST / SIGMA * Sqr(EXPIRATION), CND_TYPE) + _
            Exp(CARRY_COST * EXPIRATION) * CND_FUNC(D1_VAL, CND_TYPE))
End Select

Exit Function
ERROR_LABEL:
FLOATING_STRIKE_LOOKBACK_OPTION_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : FIXED_STRIKE_LOOKBACK_OPTION_FUNC
'DESCRIPTION   : Fixed strike lookback options
'LIBRARY       : DERIVATIVES
'GROUP         : LOOKBACK
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function FIXED_STRIKE_LOOKBACK_OPTION_FUNC(ByVal SPOT As Double, _
ByVal MIN_SPOT As Double, _
ByVal MAX_SPOT As Double, _
ByVal STRIKE As Double, _
ByVal EXPIRATION As Double, _
ByVal RATE As Double, _
ByVal CARRY_COST As Double, _
ByVal SIGMA As Double, _
Optional ByVal OPTION_FLAG As Integer = 1, _
Optional ByVal CND_TYPE As Integer = 0)
    
'------------------------------------------------------
'MIN_SPOT: Observed minimum
'MAX_SPOT: Observed Maximum
'FIRST_TENOR: First time period
'EXPIRATION: Time to maturity
'------------------------------------------------------
    
Dim D1_VAL As Double
Dim D2_VAL As Double
    
Dim E1_VAL As Double
Dim E2_VAL As Double
    
Dim TEMP_SPOT As Double
    
On Error GoTo ERROR_LABEL
    
If OPTION_FLAG <> 1 Then: OPTION_FLAG = -1
    
Select Case OPTION_FLAG
Case 1 ', "c", "call"
    TEMP_SPOT = MAX_SPOT
Case Else '-1 ', "p", "put"
    TEMP_SPOT = MIN_SPOT
End Select

D1_VAL = (Log(SPOT / STRIKE) + (CARRY_COST + SIGMA ^ 2 / 2) * EXPIRATION) / _
(SIGMA * Sqr(EXPIRATION))
    
D2_VAL = D1_VAL - SIGMA * Sqr(EXPIRATION)
    
E1_VAL = (Log(SPOT / TEMP_SPOT) + (CARRY_COST + SIGMA ^ 2 / 2) * EXPIRATION) / _
(SIGMA * Sqr(EXPIRATION))
    
E2_VAL = E1_VAL - SIGMA * Sqr(EXPIRATION)
    
Select Case OPTION_FLAG
'----------------------------------------------------------------------------------
Case 1 ', "c", "call"
'----------------------------------------------------------------------------------
    If STRIKE > TEMP_SPOT Then
        FIXED_STRIKE_LOOKBACK_OPTION_FUNC = SPOT * Exp((CARRY_COST - RATE) * EXPIRATION) * _
            CND_FUNC(D1_VAL, CND_TYPE) - STRIKE * Exp(-RATE * EXPIRATION) * _
            CND_FUNC(D2_VAL, CND_TYPE) + SPOT * Exp(-RATE _
            * EXPIRATION) * SIGMA ^ 2 / (2 * CARRY_COST) * _
            (-(SPOT / STRIKE) ^ (-2 * CARRY_COST / SIGMA ^ 2) * _
            CND_FUNC(D1_VAL - 2 * CARRY_COST / SIGMA * Sqr(EXPIRATION), CND_TYPE) + _
            Exp(CARRY_COST * EXPIRATION) * CND_FUNC(D1_VAL, CND_TYPE))
    ElseIf STRIKE <= TEMP_SPOT Then
        FIXED_STRIKE_LOOKBACK_OPTION_FUNC = Exp(-RATE * EXPIRATION) * (TEMP_SPOT - STRIKE) _
            + SPOT * Exp((CARRY_COST - RATE) * EXPIRATION) _
            * CND_FUNC(E1_VAL, CND_TYPE) - Exp(-RATE * EXPIRATION) * TEMP_SPOT * _
            CND_FUNC(E2_VAL, CND_TYPE) + _
            SPOT * Exp(-RATE * EXPIRATION) * SIGMA ^ 2 / _
            (2 * CARRY_COST) * (-(SPOT / TEMP_SPOT) ^ (-2 * CARRY_COST / _
            SIGMA ^ 2) * CND_FUNC(E1_VAL - 2 * CARRY_COST / SIGMA * _
            Sqr(EXPIRATION), CND_TYPE) _
            + Exp(CARRY_COST * EXPIRATION) * CND_FUNC(E1_VAL, CND_TYPE))
    End If
'----------------------------------------------------------------------------------
Case Else '-1 ', "p", "put"
'----------------------------------------------------------------------------------
    
    If STRIKE < TEMP_SPOT Then
        FIXED_STRIKE_LOOKBACK_OPTION_FUNC = -SPOT * Exp((CARRY_COST - RATE) * EXPIRATION) * _
            CND_FUNC(-D1_VAL, CND_TYPE) + STRIKE * Exp(-RATE * EXPIRATION) _
            * CND_FUNC(-D1_VAL + SIGMA * Sqr(EXPIRATION), CND_TYPE) + SPOT * Exp(-RATE * _
            EXPIRATION) * SIGMA ^ 2 / (2 * CARRY_COST) * ((SPOT / STRIKE) ^ _
            (-2 * CARRY_COST / SIGMA ^ 2) * CND_FUNC(-D1_VAL + 2 * CARRY_COST / SIGMA * _
            Sqr(EXPIRATION), CND_TYPE) - Exp(CARRY_COST * EXPIRATION) * _
            CND_FUNC(-D1_VAL, CND_TYPE))
    ElseIf STRIKE >= TEMP_SPOT Then
        FIXED_STRIKE_LOOKBACK_OPTION_FUNC = Exp(-RATE * EXPIRATION) * (STRIKE - TEMP_SPOT) - _
            SPOT * Exp((CARRY_COST - RATE) * EXPIRATION) * _
            CND_FUNC(-E1_VAL, CND_TYPE) + Exp(-RATE * EXPIRATION) * TEMP_SPOT * _
            CND_FUNC(-E1_VAL + SIGMA * Sqr(EXPIRATION), CND_TYPE) + Exp(-RATE * _
            EXPIRATION) * SIGMA ^ 2 / (2 * _
            CARRY_COST) * SPOT * ((SPOT / TEMP_SPOT) ^ (-2 * CARRY_COST / _
            SIGMA ^ 2) * CND_FUNC(-E1_VAL + 2 * CARRY_COST / SIGMA * Sqr(EXPIRATION), _
            CND_TYPE) - Exp(CARRY_COST * EXPIRATION) * CND_FUNC(-E1_VAL, CND_TYPE))
    End If
End Select

Exit Function
ERROR_LABEL:
FIXED_STRIKE_LOOKBACK_OPTION_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : PARTIAL_TIME_FLOATING_STRIKE_LOOKBACK_OPTION_FUNC
'DESCRIPTION   : Partial-time floating strike lookback options
'LIBRARY       : DERIVATIVES
'GROUP         : LOOKBACK
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function PARTIAL_TIME_FLOATING_STRIKE_LOOKBACK_OPTION_FUNC(ByVal SPOT As Double, _
ByVal MIN_SPOT As Double, _
ByVal MAX_SPOT As Double, _
ByVal LOOK_BACK_PERIOD As Double, _
ByVal EXPIRATION As Double, _
ByVal RATE As Double, _
ByVal CARRY_COST As Double, _
ByVal SIGMA As Double, _
ByVal LAMBDA As Double, _
Optional ByVal OPTION_FLAG As Integer = 1, _
Optional ByVal CND_TYPE As Integer = 0, _
Optional ByVal CBND_TYPE As Integer = 0)

'MIN_SPOT: Observed minimum
'MAX_SPOT: Observed Maximum
'LOOK_BACK_PERIOD: Length lookback period
'EXPIRATION: Time to maturity
'LAMBDA: Above/bellow actual extremum
  
Dim D1_VAL As Double
Dim D2_VAL As Double
    
Dim E1_VAL As Double
Dim E2_VAL As Double
    
Dim F1_VAL As Double
Dim F2_VAL As Double
    
Dim G1_VAL As Double
Dim G2_VAL As Double
    
Dim TEMP_SPOT As Double
    
Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double
Dim CTEMP_VAL As Double
    
On Error GoTo ERROR_LABEL
    
If OPTION_FLAG <> 1 Then: OPTION_FLAG = -1

Select Case OPTION_FLAG
Case 1 ', "c", "call" '---> Call on the minimum
    TEMP_SPOT = MIN_SPOT
Case Else '-1 ', "p", "put" '---> Put on the maximum
    TEMP_SPOT = MAX_SPOT
End Select
    
D1_VAL = (Log(SPOT / TEMP_SPOT) + (CARRY_COST + SIGMA ^ 2 / 2) * EXPIRATION) / _
        (SIGMA * Sqr(EXPIRATION))
D2_VAL = D1_VAL - SIGMA * Sqr(EXPIRATION)
    
E1_VAL = (CARRY_COST + SIGMA ^ 2 / 2) * (EXPIRATION - LOOK_BACK_PERIOD) / _
        (SIGMA * Sqr(EXPIRATION - LOOK_BACK_PERIOD))
E2_VAL = E1_VAL - SIGMA * Sqr(EXPIRATION - LOOK_BACK_PERIOD)
    
F1_VAL = (Log(SPOT / TEMP_SPOT) + (CARRY_COST + SIGMA ^ 2 / 2) * _
        LOOK_BACK_PERIOD) / (SIGMA * Sqr(LOOK_BACK_PERIOD))
F2_VAL = F1_VAL - SIGMA * Sqr(LOOK_BACK_PERIOD)
    
G1_VAL = Log(LAMBDA) / (SIGMA * Sqr(EXPIRATION))
G2_VAL = Log(LAMBDA) / (SIGMA * Sqr(EXPIRATION - LOOK_BACK_PERIOD))

'----------------------------------------------------------------------------------
Select Case OPTION_FLAG
'----------------------------------------------------------------------------------
Case 1 ', "c", "call"
'----------------------------------------------------------------------------------
    
        ATEMP_VAL = SPOT * Exp((CARRY_COST - RATE) * EXPIRATION) * _
            CND_FUNC(D1_VAL - G1_VAL, CND_TYPE) - LAMBDA * TEMP_SPOT * Exp(-RATE * _
            EXPIRATION) * CND_FUNC(D2_VAL - G1_VAL, CND_TYPE)
        
        BTEMP_VAL = Exp(-RATE * EXPIRATION) * SIGMA ^ 2 / (2 * CARRY_COST) * LAMBDA * _
            SPOT * ((SPOT / TEMP_SPOT) ^ (-2 * CARRY_COST / SIGMA ^ 2) * _
            CBND_FUNC(-F1_VAL + 2 * CARRY_COST * Sqr(LOOK_BACK_PERIOD) / _
            SIGMA, -D1_VAL + 2 * CARRY_COST * Sqr(EXPIRATION) / SIGMA - G1_VAL, _
            Sqr(LOOK_BACK_PERIOD / EXPIRATION), CND_TYPE, CBND_TYPE) - _
            Exp(CARRY_COST * EXPIRATION) * LAMBDA ^ _
            (2 * CARRY_COST / SIGMA ^ 2) * CBND_FUNC(-D1_VAL - G1_VAL, E1_VAL + G2_VAL, _
            -Sqr(1 - LOOK_BACK_PERIOD / EXPIRATION), CND_TYPE, CBND_TYPE)) + _
            SPOT * Exp((CARRY_COST - RATE) * EXPIRATION) * CBND_FUNC(-D1_VAL + G1_VAL, E1_VAL - G2_VAL, _
            -Sqr(1 - LOOK_BACK_PERIOD / EXPIRATION), CND_TYPE, CBND_TYPE)
        
        CTEMP_VAL = Exp(-RATE * EXPIRATION) * LAMBDA * TEMP_SPOT * CBND_FUNC(-F2_VAL, D2_VAL - G1_VAL, _
            -Sqr(LOOK_BACK_PERIOD / EXPIRATION), CND_TYPE, CBND_TYPE) - _
            Exp(-CARRY_COST * (EXPIRATION - LOOK_BACK_PERIOD)) * Exp((CARRY_COST - _
            RATE) * EXPIRATION) * (1 + SIGMA ^ 2 / (2 * CARRY_COST)) * _
            LAMBDA * SPOT * CND_FUNC(E2_VAL - G2_VAL, CND_TYPE) * CND_FUNC(-F1_VAL, CND_TYPE)
    
'----------------------------------------------------------------------------------
Case Else '-1', "p", "put"
'----------------------------------------------------------------------------------
        
        ATEMP_VAL = LAMBDA * TEMP_SPOT * Exp(-RATE * EXPIRATION) * _
            CND_FUNC(-D2_VAL + G1_VAL, CND_TYPE) - SPOT * Exp((CARRY_COST - RATE) * _
            EXPIRATION) * CND_FUNC(-D1_VAL + G1_VAL, CND_TYPE)
        
        BTEMP_VAL = -Exp(-RATE * EXPIRATION) * SIGMA ^ 2 / (2 * CARRY_COST) * LAMBDA * _
            SPOT * ((SPOT / TEMP_SPOT) ^ (-2 * CARRY_COST / SIGMA ^ 2) * CBND_FUNC(F1_VAL - _
            2 * CARRY_COST * Sqr(LOOK_BACK_PERIOD) / SIGMA, D1_VAL - 2 * CARRY_COST * _
            Sqr(EXPIRATION) / SIGMA + G1_VAL, Sqr(LOOK_BACK_PERIOD / EXPIRATION), _
            CND_TYPE, CBND_TYPE) - Exp(CARRY_COST * EXPIRATION) * LAMBDA ^ (2 * _
            CARRY_COST / SIGMA ^ 2) * CBND_FUNC(D1_VAL + G1_VAL, -E1_VAL - G2_VAL, -Sqr(1 - _
            LOOK_BACK_PERIOD / EXPIRATION), CND_TYPE, CBND_TYPE)) - _
            SPOT * Exp((CARRY_COST - RATE) * EXPIRATION) * CBND_FUNC(D1_VAL - G1_VAL, _
            -E1_VAL + G2_VAL, -Sqr(1 - LOOK_BACK_PERIOD / EXPIRATION), CND_TYPE, CBND_TYPE)
        
        CTEMP_VAL = -Exp(-RATE * EXPIRATION) * LAMBDA * TEMP_SPOT * CBND_FUNC(F2_VAL, -D2_VAL + G1_VAL, _
            -Sqr(LOOK_BACK_PERIOD / EXPIRATION), CND_TYPE, CBND_TYPE) + _
            Exp(-CARRY_COST * (EXPIRATION - LOOK_BACK_PERIOD)) * Exp((CARRY_COST - _
            RATE) * EXPIRATION) * (1 + SIGMA ^ 2 / (2 * CARRY_COST)) * _
            LAMBDA * SPOT * CND_FUNC(-E2_VAL + G2_VAL, CND_TYPE) * CND_FUNC(F1_VAL, CND_TYPE)
'----------------------------------------------------------------------------------
End Select
'----------------------------------------------------------------------------------
  
PARTIAL_TIME_FLOATING_STRIKE_LOOKBACK_OPTION_FUNC = ATEMP_VAL + BTEMP_VAL + CTEMP_VAL

Exit Function
ERROR_LABEL:
PARTIAL_TIME_FLOATING_STRIKE_LOOKBACK_OPTION_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : PARTIAL_TIME_FIXED_STRIKE_LOOKBACK_OPTION_FUNC
'DESCRIPTION   : Partial-time fixed strike lookback options
'LIBRARY       : DERIVATIVES
'GROUP         : LOOKBACK
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function PARTIAL_TIME_FIXED_STRIKE_LOOKBACK_OPTION_FUNC(ByVal SPOT As Double, _
ByVal STRIKE As Double, _
ByVal LOOK_BACK_PERIOD As Double, _
ByVal EXPIRATION As Double, _
ByVal RATE As Double, _
ByVal CARRY_COST As Double, _
ByVal SIGMA As Double, _
Optional ByVal OPTION_FLAG As Integer = 1, _
Optional ByVal CND_TYPE As Integer = 0, _
Optional ByVal CBND_TYPE As Integer = 0)

'LOOK_BACK_PERIOD: Time to start of lookback period
'EXPIRATION: Time to maturity

Dim D1_VAL As Double
Dim D2_VAL As Double
    
Dim E1_VAL As Double
Dim E2_VAL As Double
    
Dim F1_VAL As Double
Dim F2_VAL As Double
    
On Error GoTo ERROR_LABEL

If OPTION_FLAG <> 1 Then: OPTION_FLAG = -1

D1_VAL = (Log(SPOT / STRIKE) + (CARRY_COST + SIGMA ^ 2 / 2) * EXPIRATION) / _
(SIGMA * Sqr(EXPIRATION))
D2_VAL = D1_VAL - SIGMA * Sqr(EXPIRATION)
    
E1_VAL = ((CARRY_COST + SIGMA ^ 2 / 2) * (EXPIRATION - LOOK_BACK_PERIOD)) / (SIGMA * _
Sqr(EXPIRATION - LOOK_BACK_PERIOD))
E2_VAL = E1_VAL - SIGMA * Sqr(EXPIRATION - LOOK_BACK_PERIOD)
    
F1_VAL = (Log(SPOT / STRIKE) + (CARRY_COST + SIGMA ^ 2 / 2) * LOOK_BACK_PERIOD) / _
(SIGMA * Sqr(LOOK_BACK_PERIOD))
F2_VAL = F1_VAL - SIGMA * Sqr(LOOK_BACK_PERIOD)
    
Select Case OPTION_FLAG
'----------------------------------------------------------------------------------
Case 1 ', "c", "call"
'----------------------------------------------------------------------------------
        
        PARTIAL_TIME_FIXED_STRIKE_LOOKBACK_OPTION_FUNC = SPOT * Exp((CARRY_COST - RATE) * _
            EXPIRATION) * CND_FUNC(D1_VAL, CND_TYPE) - Exp(-RATE * EXPIRATION) * STRIKE * _
            CND_FUNC(D2_VAL, CND_TYPE) + SPOT * Exp(-RATE * _
            EXPIRATION) * SIGMA ^ 2 / (2 * CARRY_COST) * (-(SPOT / STRIKE) ^ (-2 * _
            CARRY_COST / SIGMA ^ 2) * CBND_FUNC(D1_VAL - 2 * CARRY_COST * Sqr(EXPIRATION) / _
            SIGMA, -F1_VAL + 2 * CARRY_COST * Sqr(LOOK_BACK_PERIOD) / SIGMA, _
            -Sqr(LOOK_BACK_PERIOD / EXPIRATION), CND_TYPE, CBND_TYPE) + _
            Exp(CARRY_COST * EXPIRATION) _
            * CBND_FUNC(E1_VAL, D1_VAL, Sqr(1 - LOOK_BACK_PERIOD / EXPIRATION), _
            CND_TYPE, CBND_TYPE)) - SPOT * Exp((CARRY_COST - RATE) * _
            EXPIRATION) * CBND_FUNC(-E1_VAL, D1_VAL, -Sqr(1 - _
            LOOK_BACK_PERIOD / EXPIRATION), CND_TYPE, CBND_TYPE) - STRIKE * _
            Exp(-RATE * EXPIRATION) * CBND_FUNC(F2_VAL, -D2_VAL, -Sqr(LOOK_BACK_PERIOD / _
            EXPIRATION), CND_TYPE, CBND_TYPE) + _
            Exp(-CARRY_COST * (EXPIRATION - LOOK_BACK_PERIOD)) * (1 - SIGMA ^ 2 / _
            (2 * CARRY_COST)) * SPOT * Exp((CARRY_COST - RATE) * EXPIRATION) * _
            CND_FUNC(F1_VAL, CND_TYPE) * CND_FUNC(-E2_VAL, CND_TYPE)
    
'----------------------------------------------------------------------------------
Case Else '-1 ', "p", "put"
'----------------------------------------------------------------------------------
    
        PARTIAL_TIME_FIXED_STRIKE_LOOKBACK_OPTION_FUNC = STRIKE * Exp(-RATE * EXPIRATION) * _
            CND_FUNC(-D2_VAL, CND_TYPE) - SPOT * Exp((CARRY_COST - RATE) * EXPIRATION) * _
            CND_FUNC(-D1_VAL, CND_TYPE) + SPOT * Exp(-RATE * EXPIRATION) * SIGMA ^ 2 _
            / (2 * CARRY_COST) * ((SPOT / STRIKE) ^ (-2 * CARRY_COST / _
            SIGMA ^ 2) * CBND_FUNC(-D1_VAL + 2 * CARRY_COST * _
            Sqr(EXPIRATION) / SIGMA, F1_VAL - 2 * CARRY_COST * Sqr(LOOK_BACK_PERIOD) / _
            SIGMA, -Sqr(LOOK_BACK_PERIOD / EXPIRATION), CND_TYPE, CBND_TYPE) - _
            Exp(CARRY_COST * EXPIRATION) * CBND_FUNC(-E1_VAL, -D1_VAL, Sqr(1 - LOOK_BACK_PERIOD / _
            EXPIRATION), CND_TYPE, CBND_TYPE)) + SPOT * Exp((CARRY_COST - RATE) * _
            EXPIRATION) * CBND_FUNC(E1_VAL, -D1_VAL, -Sqr(1 - LOOK_BACK_PERIOD / _
            EXPIRATION), CND_TYPE, CBND_TYPE) + _
            STRIKE * Exp(-RATE * EXPIRATION) * CBND_FUNC(-F2_VAL, D2_VAL, _
            -Sqr(LOOK_BACK_PERIOD / _
            EXPIRATION), CND_TYPE, CBND_TYPE) - Exp(-CARRY_COST * _
            (EXPIRATION - LOOK_BACK_PERIOD)) _
            * (1 - SIGMA ^ 2 / (2 * CARRY_COST)) * SPOT * _
            Exp((CARRY_COST - RATE) * EXPIRATION) * CND_FUNC(-F1_VAL, CND_TYPE) * _
            CND_FUNC(E2_VAL, CND_TYPE)
        
End Select

Exit Function
ERROR_LABEL:
PARTIAL_TIME_FIXED_STRIKE_LOOKBACK_OPTION_FUNC = Err.number
End Function

'---------------------------------------------------------------------------------
'---------------------------------------------------------------------------------
'---------------------------------------------------------------------------------

'About Lookback Options

'Description

'The lookback option is unique because it gives the holder the right
'to buy an asset at its lowest price or sell it at its highest price
'attained over the life of the option. At expiration you lookback and
'choose the best price that occurred during the option term.

'For a lookback call option, the lowest observed price is selected and
'is applied as the strike exercise price.

'For a lookback put option, the highest price is selected and is applied
'as the exercise price.

'The holder of a lookback option can never miss the best underlying asset
'price. These options reduce regret, since they guarantee a payoff if the
'option is in-the-money at any point during its life.

'Benefits
'Pays the largest in-the-money amount over the life of the option.
'The lookback call owner can buy at the lowest observed price or rate.
'The lookback put owner can sell at the highest observed price or rate.
'A lookback option can never be out-of-the money.
'The lookback holder gains economic value through hindsight.

'Features

'The lookback and standard call option prices converge as the underlying
'price increases.

'A lookback option is more expensive than a standard option.

'A justification for this higher cost is that they optimize market timing- a call
'option provides the best timing when awaiting an increase in the underlying asset
'price while the put gives the best timing when expecting a downturn.

'An in-the-money lookback approaches the value of the standard option.

'American lookback values can be approximated by the European model; the value
'of early exercise is low since the conditions for early exercise would cause
'the strike bonus to approach zero.

'A rule of thumb is that a at-the-money lookback option when issued will be
'priced at about two times a standard option.

'Intrinsic Value Formula
'The contract payoff at expiration is:
'Call: max{0, S_t Min(S_0, S_1,...,S_t)}
'Put: max{0, Max(S_0, S_1,...,S_t) S_t}

'Aliases
'Lookback options are also known as Mocatta options.

'Uses
'1) Consider a six-month currency lookback call option on 1 million BP
'against US dollars. At option expiration, you can lookback over the
'preceding six months and chose to accept sterling at the most favorable
'exchange rate that occurred. This guarantees the a no-regrets result since
'the best exchange rate will be achieved.

'2) A U.S. manufacturer buys raw materials from a Canadian supplier. Upon
'receipt, he has until months end to settle and is thus exposed to foreign
'exchange risk on a monthly basis. The manufacturer would like to lock-in
'the most favorable exchange rate in that monthly interval.

'Solution:
'Time = 0 mo.: FX = Can$1.34/US$1.00. Buy a European-style (exercise only at
'months end) lookback put option on US$. Payoff = max [0, Spot Price US$min]
'Time = 1 mo.:

'Case 1: FX = Can$1.32/US$1.00, and this is the minimum value the dollar reached
'in the month. The option is at-the-money.

'Case 2: FX = Can$1.33/US$1.00. US$min = Can$1.28.
'In-the-money: payoff = (1.33 1.28 P*1,LO) Can$/US$.
'P*1,LO = lookback option premium, future valued 1 month.

'3) Currency-linked bond issues, to avoid missing the best currency rates.
'4) Open-end offshore investment funds, to assure each new investor a lock-in
'of the best currency rates throughout participation in the fund.

'---------------------------------------------------------------------------------
'---------------------------------------------------------------------------------
'---------------------------------------------------------------------------------
