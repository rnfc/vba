Attribute VB_Name = "FINAN_DERIV_BS_GREEKS_LIBR"

'---------------------------------------------------------------------------------
'---------------------------------------------------------------------------------
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'---------------------------------------------------------------------------------
'---------------------------------------------------------------------------------

'************************************************************************************
'************************************************************************************
'FUNCTION      : BLACK_SCHOLES_GREEKS_TABLE_FUNC
'DESCRIPTION   : Black-Scholes Option Pricing Model & Option Greeks: Option
'Greeks measure the sensitivity of the option from its parameters
'References:
'Option Volatility and Pricing by Sheldon Natenberg
'Financial Models using Excel by Simon Benninga
'LIBRARY       : DERIVATIVES
'GROUP         : BS_GREEKS
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function BLACK_SCHOLES_GREEKS_TABLE_FUNC(ByVal ASSET_PRICE_RNG As Variant, _
ByVal STRIKE_PRICE_RNG As Variant, _
ByVal EXPIRATION_RNG As Variant, _
ByVal RISK_FREE_RATE_RNG As Variant, _
ByVal CARRY_COST_RNG As Variant, _
ByVal VOLATILITY_RNG As Variant, _
Optional ByVal OPTION_FLAG_RNG As Variant = 1, _
Optional ByVal CND_TYPE_RNG As Variant = 0)

'OPTION_FLAG = 1 --> CALL_OPTION
'OPTION_FLAG = -1 --> PUT_OPTION
'CARRY_COST = COST_OF_CARRY_COST = (RISK_FREE_RATE - DIVIDEND_RATE)

Dim j As Long
Dim NCOLUMNS As Long

Dim ASSET_PRICE As Double
Dim STRIKE_PRICE As Double
Dim EXPIRATION As Double
Dim RISK_FREE_RATE As Double
Dim CARRY_COST As Double
Dim VOLATILITY As Double
Dim OPTION_FLAG As Integer
Dim CND_TYPE As Integer

Dim D1_VAL As Double
Dim D2_VAL As Double

Dim OPTION_VAL As Double
Dim CARRY_COST_VAL As Double
Dim BANG_VAL As Double

Dim DELTA_VAL As Double
Dim GAMMA_VAL As Double
Dim THETA_VAL As Double
Dim VEGA_VAL As Double
Dim RHO_VAL As Double

Dim NORM_D1_VAL As Double
Dim NEG_NORM_D1_VAL As Double
Dim NORM_D2_VAL As Double
Dim NEG_NORM_D2_VAL As Double

Dim ASSET_PRICE_VECTOR As Variant
Dim STRIKE_PRICE_VECTOR As Variant
Dim EXPIRATION_VECTOR As Variant
Dim RISK_FREE_RATE_VECTOR As Variant
Dim CARRY_COST_VECTOR As Variant
Dim VOLATILITY_VECTOR As Variant
Dim OPTION_FLAG_VECTOR As Variant
Dim CND_TYPE_VECTOR As Variant

Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

'----------------------------------------------------------------------------
If IsArray(ASSET_PRICE_RNG) = True Then
    ASSET_PRICE_VECTOR = ASSET_PRICE_RNG
    If UBound(ASSET_PRICE_VECTOR, 2) = 1 Then
        ASSET_PRICE_VECTOR = MATRIX_TRANSPOSE_FUNC(ASSET_PRICE_VECTOR)
    End If
Else
    ReDim ASSET_PRICE_VECTOR(1 To 1, 1 To 1)
    ASSET_PRICE_VECTOR(1, 1) = ASSET_PRICE_RNG
End If
NCOLUMNS = UBound(ASSET_PRICE_VECTOR, 2)
'----------------------------------------------------------------------------
If IsArray(STRIKE_PRICE_RNG) = True Then
    STRIKE_PRICE_VECTOR = STRIKE_PRICE_RNG
    If UBound(STRIKE_PRICE_VECTOR, 2) = 1 Then
        STRIKE_PRICE_VECTOR = MATRIX_TRANSPOSE_FUNC(STRIKE_PRICE_VECTOR)
    End If
Else
    ReDim STRIKE_PRICE_VECTOR(1 To 1, 1 To 1)
    STRIKE_PRICE_VECTOR(1, 1) = STRIKE_PRICE_RNG
End If
If UBound(STRIKE_PRICE_VECTOR, 2) <> NCOLUMNS Then: GoTo ERROR_LABEL
'----------------------------------------------------------------------------
If IsArray(EXPIRATION_RNG) = True Then
    EXPIRATION_VECTOR = EXPIRATION_RNG
    If UBound(EXPIRATION_VECTOR, 2) = 1 Then
        EXPIRATION_VECTOR = MATRIX_TRANSPOSE_FUNC(EXPIRATION_VECTOR)
    End If
Else
    ReDim EXPIRATION_VECTOR(1 To 1, 1 To 1)
    EXPIRATION_VECTOR(1, 1) = EXPIRATION_RNG
End If
If UBound(EXPIRATION_VECTOR, 2) <> NCOLUMNS Then: GoTo ERROR_LABEL
'----------------------------------------------------------------------------
If IsArray(RISK_FREE_RATE_RNG) = True Then
    RISK_FREE_RATE_VECTOR = RISK_FREE_RATE_RNG
    If UBound(RISK_FREE_RATE_VECTOR, 2) = 1 Then
        RISK_FREE_RATE_VECTOR = MATRIX_TRANSPOSE_FUNC(RISK_FREE_RATE_VECTOR)
    End If
Else
    ReDim RISK_FREE_RATE_VECTOR(1 To 1, 1 To 1)
    RISK_FREE_RATE_VECTOR(1, 1) = RISK_FREE_RATE_RNG
End If
If UBound(RISK_FREE_RATE_VECTOR, 2) <> NCOLUMNS Then: GoTo ERROR_LABEL
'----------------------------------------------------------------------------
If IsArray(CARRY_COST_RNG) = True Then
    CARRY_COST_VECTOR = CARRY_COST_RNG
    If UBound(CARRY_COST_VECTOR, 2) = 1 Then
        CARRY_COST_VECTOR = MATRIX_TRANSPOSE_FUNC(CARRY_COST_VECTOR)
    End If
Else
    ReDim CARRY_COST_VECTOR(1 To 1, 1 To 1)
    CARRY_COST_VECTOR(1, 1) = CARRY_COST_RNG
End If
If UBound(CARRY_COST_VECTOR, 2) <> NCOLUMNS Then: GoTo ERROR_LABEL
'----------------------------------------------------------------------------
If IsArray(VOLATILITY_RNG) = True Then
    VOLATILITY_VECTOR = VOLATILITY_RNG
    If UBound(VOLATILITY_VECTOR, 2) = 1 Then
        VOLATILITY_VECTOR = MATRIX_TRANSPOSE_FUNC(VOLATILITY_VECTOR)
    End If
Else
    ReDim VOLATILITY_VECTOR(1 To 1, 1 To 1)
    VOLATILITY_VECTOR(1, 1) = VOLATILITY_RNG
End If
If UBound(VOLATILITY_VECTOR, 2) <> NCOLUMNS Then: GoTo ERROR_LABEL
'----------------------------------------------------------------------------
If IsArray(OPTION_FLAG_RNG) = True Then
    OPTION_FLAG_VECTOR = OPTION_FLAG_RNG
    If UBound(OPTION_FLAG_VECTOR, 2) = 1 Then
        OPTION_FLAG_VECTOR = MATRIX_TRANSPOSE_FUNC(OPTION_FLAG_VECTOR)
    End If
Else
    ReDim OPTION_FLAG_VECTOR(1 To 1, 1 To NCOLUMNS)
    For j = 1 To NCOLUMNS
        OPTION_FLAG_VECTOR(1, j) = OPTION_FLAG_RNG
    Next j
End If
'----------------------------------------------------------------------------
If IsArray(CND_TYPE_RNG) = True Then
    CND_TYPE_VECTOR = CND_TYPE_RNG
    If UBound(CND_TYPE_VECTOR, 2) = 1 Then
        CND_TYPE_VECTOR = MATRIX_TRANSPOSE_FUNC(CND_TYPE_VECTOR)
    End If
Else
    ReDim CND_TYPE_VECTOR(1 To 1, 1 To NCOLUMNS)
    For j = 1 To NCOLUMNS
        CND_TYPE_VECTOR(1, j) = CND_TYPE_RNG
    Next j
End If
'----------------------------------------------------------------------------
ReDim TEMP_MATRIX(1 To 8, 1 To NCOLUMNS + 1)
TEMP_MATRIX(1, 1) = "THEORETICAL PRICE"
TEMP_MATRIX(2, 1) = "CARRYING"
TEMP_MATRIX(3, 1) = "BANG"
TEMP_MATRIX(4, 1) = "DELTA"
TEMP_MATRIX(5, 1) = "GAMMA"
TEMP_MATRIX(6, 1) = "THETA"
TEMP_MATRIX(7, 1) = "VEGA"
TEMP_MATRIX(8, 1) = "RHO"
'----------------------------------------------------------------------------

'----------------------------------------------------------------------------
For j = 1 To NCOLUMNS
'----------------------------------------------------------------------------

    ASSET_PRICE = ASSET_PRICE_VECTOR(1, j)
    STRIKE_PRICE = STRIKE_PRICE_VECTOR(1, j)
    EXPIRATION = EXPIRATION_VECTOR(1, j)
    RISK_FREE_RATE = RISK_FREE_RATE_VECTOR(1, j)
    CARRY_COST = CARRY_COST_VECTOR(1, j)
    VOLATILITY = VOLATILITY_VECTOR(1, j)
    OPTION_FLAG = OPTION_FLAG_VECTOR(1, j)
    CND_TYPE = CND_TYPE_VECTOR(1, j)

    D1_VAL = (Log(ASSET_PRICE / STRIKE_PRICE) + (CARRY_COST + VOLATILITY ^ 2 / 2) * EXPIRATION) / _
    (VOLATILITY * Sqr(EXPIRATION))

    D2_VAL = D1_VAL - VOLATILITY * Sqr(EXPIRATION)
    
    NORM_D1_VAL = CND_FUNC(D1_VAL, CND_TYPE)
    NEG_NORM_D1_VAL = CND_FUNC(-D1_VAL, CND_TYPE)
    
    NORM_D2_VAL = CND_FUNC(D2_VAL, CND_TYPE)
    NEG_NORM_D2_VAL = CND_FUNC(-D2_VAL, CND_TYPE)
    
'---------------------------------------------------------------------------
    Select Case OPTION_FLAG
'---------------------------------------------------------------------------
    Case 1 ', "c", "CALL"
'---------------------------------------------------------------------------
            
        OPTION_VAL = ASSET_PRICE * Exp((CARRY_COST - RISK_FREE_RATE) * EXPIRATION) * NORM_D1_VAL _
        - STRIKE_PRICE * Exp(-RISK_FREE_RATE * EXPIRATION) * NORM_D2_VAL
        'Generalized Black and Scholes
        
        CARRY_COST_VAL = EXPIRATION * ASSET_PRICE * Exp((CARRY_COST - RISK_FREE_RATE) * _
                     EXPIRATION) * NORM_D1_VAL
            
        DELTA_VAL = Exp((CARRY_COST - RISK_FREE_RATE) * EXPIRATION) * NORM_D1_VAL
        'Delta for the generalized Black and Scholes formula
          
        BANG_VAL = DELTA_VAL * ASSET_PRICE / OPTION_VAL 'BANG FOR THE BUCK
        
        'Delta measures RISK_FREE_RATE of change of the option's value with respect to
        'the stock pricing; it is the first differential of option price with
        'respect to the price of the underlying asset.  Delta also changes
        'gradually over time even if there is no price movement of the
        'underlying asset.  The change in delta for a given change in the
        'asset price is known as Gamma.
            
        GAMMA_VAL = Exp((CARRY_COST - RISK_FREE_RATE) * EXPIRATION) * _
        NORMAL_MASS_DIST_FUNC(D1_VAL) / _
        (ASSET_PRICE * VOLATILITY * Sqr(EXPIRATION))
        'Gamma for the generalized Black and Scholes formula
        
        'Gamma is the second derivative of the option value with respect
        'to the price of the underlying asset.  Variation in Delta requires
        'that a hedged position be rebalanced if it is to remain delta
        'neutral after the price of the underlying asset has changed.
        'How much adjustment needed depends on how much the Delta
        'changes, that is, on Gamma.
            
        THETA_VAL = -ASSET_PRICE * Exp((CARRY_COST - RISK_FREE_RATE) * EXPIRATION) * _
        NORMAL_MASS_DIST_FUNC(D1_VAL) * VOLATILITY / _
        (2 * Sqr(EXPIRATION)) - (CARRY_COST - RISK_FREE_RATE) * ASSET_PRICE * _
        Exp((CARRY_COST - RISK_FREE_RATE) * EXPIRATION) _
        * NORM_D1_VAL - RISK_FREE_RATE * STRIKE_PRICE * Exp(-RISK_FREE_RATE * EXPIRATION) * NORM_D2_VAL
        'Theta for the generalized Black and Scholes formula
    
        'Theta refers to the RISK_FREE_RATE of time decay for an option.
        'It is the first differential of the option value with
        'respect to time.  Holding all other things constant, an
        'option loses value as it approaching to the expiration
        'day.  Theta measures the cost of holding an option long,
        'and the reward for writing it.
            
        VEGA_VAL = ASSET_PRICE * Exp((CARRY_COST - RISK_FREE_RATE) * EXPIRATION) * _
        NORMAL_MASS_DIST_FUNC(D1_VAL) * Sqr(EXPIRATION)
        'Vega for the generalized Black and Scholes formula
        
        'Vega measures the relationship between the volatility of the
        'underlying asset and the option value.  It is the first
        'differential of the option price with respect
        'to the volatility (standard deviation).  The more volatility
        'the underlying asset is, the more valuable the option becomes
        'since the chance for the option to be deep-in-the-money is greater.
    
        
        If CARRY_COST <> 0 Then 'Rho for the generalized Black and Scholes formula
            RHO_VAL = EXPIRATION * STRIKE_PRICE * Exp(-RISK_FREE_RATE * EXPIRATION) * _
                       NORM_D2_VAL
        Else
            RHO_VAL = -EXPIRATION * OPTION_VAL
        End If
    
        'Rho measures the sensitivity of the option value to the interest
        'rate.  It is the derivative of the option value with respect to
        'the interest rate.  The higher the interest rate, the greater the
        'time value of the option.  Hence, Rho is positive for calls and
        'negative for puts.  For both calls and puts, the longer the time
        'to expiration, the larger is the effect of the interest rate
        'on the option value.
    
'---------------------------------------------------------------------------
    Case Else '-1 ', "p", "PUT"
'---------------------------------------------------------------------------
        
        OPTION_VAL = STRIKE_PRICE * Exp(-RISK_FREE_RATE * EXPIRATION) * NEG_NORM_D2_VAL - ASSET_PRICE _
        * Exp((CARRY_COST - RISK_FREE_RATE) * EXPIRATION) * NEG_NORM_D1_VAL
        'Generalized Black and Scholes
        
        CARRY_COST_VAL = -EXPIRATION * ASSET_PRICE * Exp((CARRY_COST - RISK_FREE_RATE) * _
                     EXPIRATION) * NEG_NORM_D1_VAL
        
        DELTA_VAL = Exp((CARRY_COST - RISK_FREE_RATE) * EXPIRATION) * (NORM_D1_VAL - 1)
        'Delta for the generalized Black and Scholes formula
    
        BANG_VAL = DELTA_VAL * ASSET_PRICE / OPTION_VAL 'BANG FOR THE BUCK
    
        GAMMA_VAL = Exp((CARRY_COST - RISK_FREE_RATE) * EXPIRATION) * _
        NORMAL_MASS_DIST_FUNC(D1_VAL) / _
        (ASSET_PRICE * VOLATILITY * Sqr(EXPIRATION))
        'Gamma for the generalized Black and Scholes formula
    
        THETA_VAL = -ASSET_PRICE * Exp((CARRY_COST - RISK_FREE_RATE) * EXPIRATION) * _
        NORMAL_MASS_DIST_FUNC(D1_VAL) * VOLATILITY / _
        (2 * Sqr(EXPIRATION)) + (CARRY_COST - RISK_FREE_RATE) * ASSET_PRICE * _
        Exp((CARRY_COST - RISK_FREE_RATE) * EXPIRATION) _
        * NEG_NORM_D1_VAL + RISK_FREE_RATE * STRIKE_PRICE * Exp(-RISK_FREE_RATE * _
        EXPIRATION) * NEG_NORM_D2_VAL _
        'Theta for the generalized Black and Scholes formula
    
        VEGA_VAL = ASSET_PRICE * Exp((CARRY_COST - RISK_FREE_RATE) * EXPIRATION) * _
        NORMAL_MASS_DIST_FUNC(D1_VAL) * Sqr(EXPIRATION)
        'Vega for the generalized Black and Scholes formula
     
        If CARRY_COST <> 0 Then 'Rho for the generalized Black and Scholes formula
            RHO_VAL = -EXPIRATION * STRIKE_PRICE * Exp(-RISK_FREE_RATE * EXPIRATION) * _
                        NEG_NORM_D2_VAL
        Else
            RHO_VAL = -EXPIRATION * OPTION_VAL
        End If
'---------------------------------------------------------------------------
    End Select
'---------------------------------------------------------------------------
        
    TEMP_MATRIX(1, j + 1) = OPTION_VAL
    TEMP_MATRIX(2, j + 1) = CARRY_COST_VAL
    TEMP_MATRIX(3, j + 1) = BANG_VAL
    TEMP_MATRIX(4, j + 1) = DELTA_VAL
    TEMP_MATRIX(5, j + 1) = GAMMA_VAL
    TEMP_MATRIX(6, j + 1) = THETA_VAL
    TEMP_MATRIX(7, j + 1) = VEGA_VAL
    TEMP_MATRIX(8, j + 1) = RHO_VAL
'----------------------------------------------------------------------------
Next j
'----------------------------------------------------------------------------

BLACK_SCHOLES_GREEKS_TABLE_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
BLACK_SCHOLES_GREEKS_TABLE_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_OPTION_DELTA_FUNC

'DESCRIPTION   : Delta measures rate of change of the option's value with
'respect to the stock pricing; it is the first differential of option price
'with respect to the price of the underlying asset.  Delta also changes
'gradually over time even if there is no price movement of the underlying
'asset.  The change in delta for a given change in the asset price is
'known as Gamma.

'LIBRARY       : DERIVATIVES
'GROUP         : BS_GREEKS
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function CALL_OPTION_DELTA_FUNC(ByVal SPOT As Double, _
ByVal STRIKE As Double, _
ByVal EXPIRATION As Double, _
ByVal RATE As Double, _
ByVal DIVD As Double, _
ByVal SIGMA As Double, _
Optional ByVal CND_TYPE As Integer = 0)

On Error GoTo ERROR_LABEL
 
CALL_OPTION_DELTA_FUNC = Exp(-DIVD * EXPIRATION) * _
    CND_FUNC(EUROPEAN_D1_DENSITY_FUNC(SPOT, _
    STRIKE, EXPIRATION, RATE, DIVD, SIGMA), CND_TYPE)

Exit Function
ERROR_LABEL:
CALL_OPTION_DELTA_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************



'************************************************************************************
'************************************************************************************
'FUNCTION      : PUT_OPTION_DELTA_FUNC

'DESCRIPTION   : Delta measures rate of change of the option's value with
'respect to the stock pricing; it is the first differential of option price
'with respect to the price of the underlying asset.  Delta also changes
'gradually over time even if there is no price movement of the underlying
'asset.  The change in delta for a given change in the asset price is
'known as Gamma.

'LIBRARY       : DERIVATIVES
'GROUP         : BS_GREEKS
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************
'************************************************************************************

Function PUT_OPTION_DELTA_FUNC(ByVal SPOT As Double, _
ByVal STRIKE As Double, _
ByVal EXPIRATION As Double, _
ByVal RATE As Double, _
ByVal DIVD As Double, _
ByVal SIGMA As Double, _
Optional ByVal CND_TYPE As Integer = 0)

On Error GoTo ERROR_LABEL
 
PUT_OPTION_DELTA_FUNC = Exp(-DIVD * EXPIRATION) * _
    (CND_FUNC(EUROPEAN_D1_DENSITY_FUNC(SPOT, _
    STRIKE, EXPIRATION, RATE, DIVD, SIGMA), CND_TYPE) - 1)

Exit Function
ERROR_LABEL:
PUT_OPTION_DELTA_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_PUT_OPTION_GAMMA_FUNC

'DESCRIPTION   : Gamma is the second derivative of the option value with
'respect to the price of the underlying asset. Variation in Delta requires
'that a hedged position be rebalanced if it is to remain delta neutral after
'the price of the underlying asset has changed.  How much adjustment needed
'depends on how much the Delta changes, that is, on Gamma.

'LIBRARY       : DERIVATIVES
'GROUP         : BS_GREEKS
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function CALL_PUT_OPTION_GAMMA_FUNC(ByVal SPOT As Double, _
ByVal STRIKE As Double, _
ByVal EXPIRATION As Double, _
ByVal RATE As Double, _
ByVal DIVD As Double, _
ByVal SIGMA As Double)
 
On Error GoTo ERROR_LABEL
 
CALL_PUT_OPTION_GAMMA_FUNC = Exp((-DIVD) * EXPIRATION) * _
    NORMAL_MASS_DIST_FUNC(EUROPEAN_D1_DENSITY_FUNC(SPOT, _
    STRIKE, EXPIRATION, RATE, DIVD, SIGMA)) / (SPOT * SIGMA * Sqr(EXPIRATION))

Exit Function
ERROR_LABEL:
CALL_PUT_OPTION_GAMMA_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_OPTION_RHO_FUNC
'DESCRIPTION   : Rho measures the sensitivity of the option value to the
'interest rate.  It is the derivative of the option value with respect to
'the interest rate.  The higher the interest rate, the greater the time
'value of the option.  Hence, Rho is positive for calls and negative for
'puts.  For both calls and puts, the longer the time to expiration, the
'larger is the effect of the interest rate on the option value.
'LIBRARY       : DERIVATIVES
'GROUP         : BS_GREEKS
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function CALL_OPTION_RHO_FUNC(ByVal SPOT As Double, _
ByVal STRIKE As Double, _
ByVal EXPIRATION As Double, _
ByVal RATE As Double, _
ByVal DIVD As Double, _
ByVal SIGMA As Double, _
Optional ByVal CND_TYPE As Integer = 0)
  
On Error GoTo ERROR_LABEL
 
    If (RATE - DIVD) <> 0 Then 'Rho for the generalized Black and Scholes formula
         CALL_OPTION_RHO_FUNC = EXPIRATION * STRIKE * Exp(-RATE * EXPIRATION) * _
         CND_FUNC(EUROPEAN_D2_DENSITY_FUNC(SPOT, STRIKE, _
         EXPIRATION, RATE, DIVD, SIGMA), CND_TYPE)
    Else
         CALL_OPTION_RHO_FUNC = -EXPIRATION * EUROPEAN_CALL_OPTION_FUNC(SPOT, _
         STRIKE, EXPIRATION, RATE, DIVD, SIGMA, CND_TYPE)
    End If
    
Exit Function
ERROR_LABEL:
CALL_OPTION_RHO_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : PUT_OPTION_RHO_FUNC
'DESCRIPTION   : Rho measures the sensitivity of the option value to the
'interest rate.  It is the derivative of the option value with respect to
'the interest rate.  The higher the interest rate, the greater the time
'value of the option.  Hence, Rho is positive for calls and negative for
'puts.  For both calls and puts, the longer the time to expiration, the
'larger is the effect of the interest rate on the option value.
'LIBRARY       : DERIVATIVES
'GROUP         : BS_GREEKS
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************
    
Function PUT_OPTION_RHO_FUNC(ByVal SPOT As Double, _
ByVal STRIKE As Double, _
ByVal EXPIRATION As Double, _
ByVal RATE As Double, _
ByVal DIVD As Double, _
ByVal SIGMA As Double, _
Optional ByVal CND_TYPE As Integer = 0)

On Error GoTo ERROR_LABEL
 
    If (RATE - DIVD) <> 0 Then 'Rho for the generalized Black and Scholes formula
         PUT_OPTION_RHO_FUNC = -EXPIRATION * STRIKE * Exp(-RATE * EXPIRATION) * _
            CND_FUNC(-EUROPEAN_D2_DENSITY_FUNC(SPOT, _
            STRIKE, EXPIRATION, RATE, DIVD, SIGMA), CND_TYPE)
    Else
         PUT_OPTION_RHO_FUNC = -EXPIRATION * EUROPEAN_PUT_OPTION_FUNC(SPOT, _
            STRIKE, EXPIRATION, _
            RATE, DIVD, SIGMA, CND_TYPE)
    End If

Exit Function
ERROR_LABEL:
PUT_OPTION_RHO_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_OPTION_THETA_FUNC
'DESCRIPTION   : Theta refers to the rate of time decay for an option.  It is
'the first differential of the option value with respect to time.  Holding all
'other things constant, an option loses value as it approaching to the
'expiration day.  Theta measures the cost of holding an option long, and the
'reward fo writing it.
'LIBRARY       : DERIVATIVES
'GROUP         : BS_GREEKS
'ID            : 007
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function CALL_OPTION_THETA_FUNC(ByVal SPOT As Double, _
ByVal STRIKE As Double, _
ByVal EXPIRATION As Double, _
ByVal RATE As Double, _
ByVal DIVD As Double, _
ByVal SIGMA As Double, _
Optional ByVal CND_TYPE As Integer = 0)
 
On Error GoTo ERROR_LABEL

CALL_OPTION_THETA_FUNC = -SPOT * Exp((-DIVD) * EXPIRATION) * _
    NORMAL_MASS_DIST_FUNC(EUROPEAN_D1_DENSITY_FUNC(SPOT, _
    STRIKE, EXPIRATION, RATE, DIVD, SIGMA)) * SIGMA / _
    (2 * Sqr(EXPIRATION)) + DIVD * SPOT * Exp((-DIVD) * EXPIRATION) _
    * CND_FUNC(EUROPEAN_D1_DENSITY_FUNC(SPOT, STRIKE, _
    EXPIRATION, RATE, DIVD, SIGMA), CND_TYPE) _
    - RATE * STRIKE * Exp(-RATE * EXPIRATION) * _
    CND_FUNC(EUROPEAN_D2_DENSITY_FUNC(SPOT, STRIKE, EXPIRATION, _
    RATE, DIVD, SIGMA), CND_TYPE)
    
Exit Function
ERROR_LABEL:
CALL_OPTION_THETA_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : PUT_OPTION_THETA_FUNC
'DESCRIPTION   : Theta refers to the rate of time decay for an option.  It is
'the first differential of the option value with respect to time.  Holding all
'other things constant, an option loses value as it approaching to the
'expiration day.  Theta measures the cost of holding an option long, and the
'reward fo writing it.
'LIBRARY       : DERIVATIVES
'GROUP         : BS_GREEKS
'ID            : 008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function PUT_OPTION_THETA_FUNC(ByVal SPOT As Double, _
ByVal STRIKE As Double, _
ByVal EXPIRATION As Double, _
ByVal RATE As Double, _
ByVal DIVD As Double, _
ByVal SIGMA As Double, _
Optional ByVal CND_TYPE As Integer = 0)
 
On Error GoTo ERROR_LABEL

PUT_OPTION_THETA_FUNC = -SPOT * Exp(-DIVD * EXPIRATION) * _
    NORMAL_MASS_DIST_FUNC(EUROPEAN_D1_DENSITY_FUNC(SPOT, _
    STRIKE, EXPIRATION, RATE, DIVD, SIGMA)) * SIGMA / _
    (2 * Sqr(EXPIRATION)) - DIVD * SPOT * Exp(-DIVD * EXPIRATION) _
    * CND_FUNC(-EUROPEAN_D1_DENSITY_FUNC(SPOT, _
    STRIKE, EXPIRATION, RATE, DIVD, SIGMA), CND_TYPE) + RATE * _
    STRIKE * Exp(-RATE * EXPIRATION) * _
    CND_FUNC(-EUROPEAN_D2_DENSITY_FUNC(SPOT, STRIKE, _
    EXPIRATION, RATE, DIVD, SIGMA), CND_TYPE)

'-((S*Nx_d1*Vol*EXP(-Div_Yield*BS_Time))/(2*SQRT(BS_Time)))-Div_Yield*S*NORMSDIST(-d_1)*EXP(-Div_Yield*BS_Time)+(Int*X*(EXP(-Int*BS_Time))*NORMSDIST(-d_2))

Exit Function
ERROR_LABEL:
PUT_OPTION_THETA_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_PUT_OPTION_VEGA_FUNC

'DESCRIPTION   : Vega measures the relationship between the volatility of the
'underlying asset and the option value.  It is the first differential of the
'option price with respect to the volatility (standard deviation).  The more
'volatility the underlying asset is, the more valuable the option becomes
'since the chance for the option to be deep-in-the-money is greater.

'LIBRARY       : DERIVATIVES
'GROUP         : BS_GREEKS
'ID            : 009
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function CALL_PUT_OPTION_VEGA_FUNC(ByVal SPOT As Double, _
ByVal STRIKE As Double, _
ByVal EXPIRATION As Double, _
ByVal RATE As Double, _
ByVal DIVD As Double, _
ByVal SIGMA As Double)
 
On Error GoTo ERROR_LABEL

CALL_PUT_OPTION_VEGA_FUNC = SPOT * Exp((-DIVD) * EXPIRATION) * _
NORMAL_MASS_DIST_FUNC(EUROPEAN_D1_DENSITY_FUNC(SPOT, _
STRIKE, EXPIRATION, RATE, DIVD, SIGMA)) * Sqr(EXPIRATION)
'=(S*SQRT(BS_Time)*Nx_d1*EXP(-Div_Yield*BS_Time))/100
Exit Function
ERROR_LABEL:
CALL_PUT_OPTION_VEGA_FUNC = Err.number
End Function
