Attribute VB_Name = "FINAN_DERIV_MIN_MAX_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : TWO_RISKY_ASSETS_MINIMUM_RAINBOW_CALL_OPTION_FUNC

'DESCRIPTION   : Call on minimum rainbow function
'An investment manager holds two negatively correlated risky assets.
'The manager expects a price movement but does not know which asset
'will increase in price and which will decrease. A put on the worse
'of the two assets will provide insurance no matter which price
'movement occurs.

'LIBRARY       : DERIVATIVES
'GROUP         : MIN_MAX
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'**********************************************************************************
'**********************************************************************************

Function TWO_RISKY_ASSETS_MINIMUM_RAINBOW_CALL_OPTION_FUNC(ByVal SPOT_A As Double, _
ByVal SPOT_B As Double, _
ByVal STRIKE As Double, _
ByVal EXPIRATION As Double, _
ByVal RATE As Double, _
ByVal QUANTITY_A As Double, _
ByVal QUANTITY_B As Double, _
ByVal DIVD_A As Double, _
ByVal DIVD_B As Double, _
ByVal SIGMA_A As Double, _
ByVal SIGMA_B As Double, _
ByVal RHO As Double, _
Optional ByVal CND_TYPE As Integer = 0, _
Optional ByVal CBND_TYPE As Integer = 0)
  
'The term rainbow option is applied to an entire class of options
'which are written on more than one underlying asset. Rainbow
'options are usually calls or puts on the best or worst of n
'underlying assets, or options which pay the best or worst of n
'assets. Spread options are a special case of rainbow options.

'Rainbow options are excellent tools for hedging the risk of
'multiple assets.

'Rainbow options provide effective hedging of assets with negative
'correlation.

'A put on the worse of two assets provides protection from price
'movements in either direction.

'Options on two highly correlated assets are less expensive than options on
'two assets which are not correlated, as lower correlation implies more
'variability in the individual prices.

'Rainbow options at exercise may deliver either the best or worse asset in
'the rainbow or a call or put option on the better or worse of the assets.
'Multi-color rainbow options could deliver the best or worst m of
'the n assets.

'The contract payoff at expiration is:
'Call on max of 2 assets: max(0, max(S1,S2) E)
'Call on min of 2 assets: max(0, min(S1,S2) E)
'Put on max of 2 assets: max(0, E max(S1,S2))
'Put on min of 2 assets: max(0, E min(S1,S2))

'Worse of 2 assets: min(S1,S2)
'Better of 2 assets: max(S1,S2)
  
  
  Dim ATEMP_VAL As Double
  Dim BTEMP_VAL As Double
  Dim CTEMP_VAL As Double
  
  Dim TEMP_SIGMA As Double
  
  On Error GoTo ERROR_LABEL
  
  ATEMP_VAL = (Log(QUANTITY_A * SPOT_A / STRIKE) + _
    (RATE - DIVD_A - 0.5 * SIGMA_A ^ 2) _
     * EXPIRATION) / (SIGMA_A * Sqr(EXPIRATION))
  
  BTEMP_VAL = (Log(QUANTITY_B * SPOT_B / STRIKE) + _
        (RATE - DIVD_B - 0.5 * SIGMA_B ^ 2) _
      * EXPIRATION) / (SIGMA_B * Sqr(EXPIRATION))
  
  TEMP_SIGMA = Sqr(SIGMA_A ^ 2 + SIGMA_B ^ 2 - _
                2 * RHO * SIGMA_A * SIGMA_B)
  
  CTEMP_VAL = QUANTITY_A * SPOT_A * Exp(-DIVD_A * EXPIRATION) * _
      CBND_FUNC(ATEMP_VAL + SIGMA_A * Sqr(EXPIRATION), _
      (Log((QUANTITY_B * SPOT_B) / (QUANTITY_A * SPOT_A)) + _
      (DIVD_A - DIVD_B - 0.5 * _
      TEMP_SIGMA ^ 2) * EXPIRATION) / (TEMP_SIGMA * Sqr(EXPIRATION)), _
      (RHO * SIGMA_B - SIGMA_A) / TEMP_SIGMA, CND_TYPE, CBND_TYPE)
  
  CTEMP_VAL = CTEMP_VAL + QUANTITY_B * SPOT_B * Exp(-DIVD_B * _
      EXPIRATION) * CBND_FUNC(BTEMP_VAL + SIGMA_B * Sqr(EXPIRATION), _
      (Log((QUANTITY_A * SPOT_A) / (QUANTITY_B * SPOT_B)) + _
      (DIVD_B - DIVD_A - 0.5 * TEMP_SIGMA ^ 2) * EXPIRATION) / _
      (TEMP_SIGMA * Sqr(EXPIRATION)), _
      (RHO * SIGMA_A - SIGMA_B) / TEMP_SIGMA, CND_TYPE, CBND_TYPE)
  
  CTEMP_VAL = CTEMP_VAL - STRIKE * Exp(-RATE * EXPIRATION) * _
      CBND_FUNC(ATEMP_VAL, BTEMP_VAL, RHO, CND_TYPE, CBND_TYPE)

  TWO_RISKY_ASSETS_MINIMUM_RAINBOW_CALL_OPTION_FUNC = CTEMP_VAL

Exit Function
ERROR_LABEL:
TWO_RISKY_ASSETS_MINIMUM_RAINBOW_CALL_OPTION_FUNC = Err.number
End Function

'**********************************************************************************
'**********************************************************************************
'FUNCTION      : TWO_RISKY_ASSETS_MAX_MIN_OPTION_FUNC
'DESCRIPTION   : Options on the maximum or the minimum of two risky assets
'LIBRARY       : DERIVATIVES
'GROUP         : MIN_MAX
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'**********************************************************************************
'**********************************************************************************

Function TWO_RISKY_ASSETS_MAX_MIN_OPTION_FUNC(ByVal SPOT_A As Double, _
ByVal SPOT_B As Double, _
ByVal STRIKE As Double, _
ByVal EXPIRATION As Double, _
ByVal RATE As Double, _
ByVal CARRY_COST_A As Double, _
ByVal CARRY_COST_B As Double, _
ByVal SIGMA_A As Double, _
ByVal SIGMA_B As Double, _
ByVal RHO As Double, _
Optional ByVal OPTION_FLAG As Integer = 1, _
Optional ByVal CND_TYPE As Integer = 0, _
Optional ByVal CBND_TYPE As Integer = 0)

'RHO = Correlation(A,B)
'[OPTION_FLAG: 1] Call on the minimum
'[OPTION_FLAG: 2] Call on the maximum
'[OPTION_FLAG: 3] Put on the minimum
'[OPTION_FLAG: 4] Put on the maximum

Dim Y1_VAL As Double
Dim Y2_VAL As Double
    
Dim RHO1_VAL As Double
Dim RHO2_VAL As Double
    
Dim TEMP_VAL As Double
Dim TEMP_SIGMA As Double
    
On Error GoTo ERROR_LABEL
   
TEMP_SIGMA = Sqr(SIGMA_A ^ 2 + SIGMA_B ^ 2 - 2 * _
             RHO * SIGMA_A * SIGMA_B)
   
RHO1_VAL = (SIGMA_A - RHO * SIGMA_B) / TEMP_SIGMA
RHO2_VAL = (SIGMA_B - RHO * SIGMA_A) / TEMP_SIGMA
    
TEMP_VAL = (Log(SPOT_A / SPOT_B) + (CARRY_COST_A - _
    CARRY_COST_B + TEMP_SIGMA ^ 2 / 2) * _
    EXPIRATION) / (TEMP_SIGMA * Sqr(EXPIRATION))
    
Y1_VAL = (Log(SPOT_A / STRIKE) + (CARRY_COST_A + SIGMA_A ^ 2 / 2) * _
    EXPIRATION) / (SIGMA_A * Sqr(EXPIRATION))
    
Y2_VAL = (Log(SPOT_B / STRIKE) + (CARRY_COST_B + SIGMA_B ^ 2 / 2) * _
    EXPIRATION) / (SIGMA_B * Sqr(EXPIRATION))
  
Select Case OPTION_FLAG
    Case 1 ', "cmin"
        TWO_RISKY_ASSETS_MAX_MIN_OPTION_FUNC = SPOT_A * Exp((CARRY_COST_A - RATE) * EXPIRATION) _
            * CBND_FUNC(Y1_VAL, -TEMP_VAL, -RHO1_VAL, CND_TYPE, CBND_TYPE) + _
            SPOT_B * Exp((CARRY_COST_B - RATE) * EXPIRATION) _
            * CBND_FUNC(Y2_VAL, TEMP_VAL - TEMP_SIGMA * Sqr(EXPIRATION), _
            -RHO2_VAL, CND_TYPE, CBND_TYPE) - STRIKE * Exp(-RATE * _
            EXPIRATION) * CBND_FUNC(Y1_VAL - SIGMA_A * Sqr(EXPIRATION), _
            Y2_VAL - SIGMA_B * Sqr(EXPIRATION), RHO, CND_TYPE, CBND_TYPE)
    Case 2 ', "cmax"
        TWO_RISKY_ASSETS_MAX_MIN_OPTION_FUNC = SPOT_A * Exp((CARRY_COST_A - RATE) * EXPIRATION) * _
            CBND_FUNC(Y1_VAL, TEMP_VAL, RHO1_VAL, CND_TYPE, CBND_TYPE) + SPOT_B * _
            Exp((CARRY_COST_B - RATE) * EXPIRATION) * _
            CBND_FUNC(Y2_VAL, -TEMP_VAL + TEMP_SIGMA * Sqr(EXPIRATION), _
            RHO2_VAL, CND_TYPE, CBND_TYPE) - STRIKE * Exp(-RATE * _
            EXPIRATION) * (1 - CBND_FUNC(-Y1_VAL + SIGMA_A * Sqr(EXPIRATION), _
            -Y2_VAL + SIGMA_B * Sqr(EXPIRATION), RHO, CND_TYPE, CBND_TYPE))
    Case 3 ', "pmin"
        TWO_RISKY_ASSETS_MAX_MIN_OPTION_FUNC = STRIKE * Exp(-RATE * EXPIRATION) - SPOT_A * _
            Exp((CARRY_COST_A - RATE) * EXPIRATION) + _
            EXCHANGE_ONE_ASSET_OPTION_FUNC(SPOT_A, SPOT_B, 1, 1, EXPIRATION, _
            RATE, CARRY_COST_A, CARRY_COST_B, SIGMA_A, _
            SIGMA_B, RHO, 0, CND_TYPE) + TWO_RISKY_ASSETS_MAX_MIN_OPTION_FUNC(SPOT_A, _
            SPOT_B, STRIKE, EXPIRATION, RATE, CARRY_COST_A, _
            CARRY_COST_B, SIGMA_A, _
            SIGMA_B, RHO, 1, CND_TYPE, CBND_TYPE)
    Case Else '4 ', "pmax"
        TWO_RISKY_ASSETS_MAX_MIN_OPTION_FUNC = STRIKE * Exp(-RATE * EXPIRATION) - SPOT_B * _
            Exp((CARRY_COST_B - RATE) * EXPIRATION) - _
            EXCHANGE_ONE_ASSET_OPTION_FUNC(SPOT_A, _
            SPOT_B, 1, 1, EXPIRATION, RATE, CARRY_COST_A, _
            CARRY_COST_B, SIGMA_A, _
            SIGMA_B, RHO, 0, CND_TYPE) + TWO_RISKY_ASSETS_MAX_MIN_OPTION_FUNC(SPOT_A, _
            SPOT_B, STRIKE, EXPIRATION, RATE, CARRY_COST_A, _
            CARRY_COST_B, SIGMA_A, _
            SIGMA_B, RHO, 2, CND_TYPE, CBND_TYPE)
End Select

Exit Function
ERROR_LABEL:
TWO_RISKY_ASSETS_MAX_MIN_OPTION_FUNC = Err.number
End Function
