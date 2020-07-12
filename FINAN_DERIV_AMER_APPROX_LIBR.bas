Attribute VB_Name = "FINAN_DERIV_AMER_APPROX_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : APPROXIMATION_AMERICAN_OPTION_FUNC
'DESCRIPTION   : AMERICAN_OPTION_APPROXIMATION
'LIBRARY       : DERIVATIVES
'GROUP         : AMER_APPROX
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function APPROXIMATION_AMERICAN_OPTION_FUNC(ByVal SPOT As Double, _
ByVal STRIKE As Double, _
ByVal EXPIRATION As Double, _
ByVal RISK_FREE_RATE As Double, _
ByVal CARRY_COST As Double, _
ByVal VOLATILITY As Double, _
Optional ByVal OPTION_FLAG As Integer = 1, _
Optional ByVal VERSION As Integer = 0, _
Optional ByVal CND_TYPE As Integer = 0, _
Optional ByVal CBND_TYPE As Integer = 0)

'-----------------------------------------------------------------------------
'Typically European equity options are priced using the Black-Scholes model
'(Black & Scholes 1973) or that model adjusted for dividends by calculating
'a continuous dividend yield. This has the effect of spreading the dividend
'payment throughout the life of the option. In the case of several dividend
'payments, this is a satisfactory solution, for example, where the option is
'on an index (where the index is paying out several dividends, spread out
'through the period of optionality).

'Thus, for European equity options where the underlying has no or several
'dividends, we will use the Black-Scholes formula. For American equity options
'with the underlying having no or several dividends, we may argue similarly.
'Here the approximation of Barone-Adesi and Whaley (Barone-Adesi & Whaley 1987)
'is popular, but i prefer the method of Bjerksund and Stensland (Bjerksund &
'Stensland 1993), (Bjerksund & Stensland 2002) as it is computationally far
'superior, and has been shown to be more accurate in long dated options.
'(Bjerksund & Stensland 2002) is a recent improvement over
'(Bjerksund & Stensland 1993).
'-----------------------------------------------------------------------------

Dim I1_VAL As Double
Dim I2_VAL As Double

Dim A_VAL As Double
Dim D_VAL As Double

Dim I_VAL As Double
Dim K_VAL As Double
Dim N_VAL As Double

Dim Q_VAL As Double
Dim SK_VAL As Double
Dim T_VAL As Double

Dim BETAT_VAL As Double
Dim BETA0_VAL As Double
Dim BETAI_VAL As Double

Dim HT_VAL As Double
Dim HT1_VAL As Double
Dim HT2_VAL As Double

Dim ALPHAT_VAL As Double
Dim ALPHA1_VAL As Double
Dim ALPHA2_VAL As Double

Dim TEMP1_VAL As Double
Dim TEMP2_VAL As Double
Dim TEMP3_VAL As Double

On Error GoTo ERROR_LABEL

'------------------------------------------------------------------------------
Select Case VERSION
'------------------------------------------------------------------------------
Case 0 'Bjerksund and Stensland American Approximation (1993)
'------------------------------------------------------------------------------
    If OPTION_FLAG = -1 Then
        TEMP1_VAL = SPOT
        TEMP2_VAL = STRIKE
        STRIKE = TEMP1_VAL
        SPOT = TEMP2_VAL
        RISK_FREE_RATE = RISK_FREE_RATE - CARRY_COST
        CARRY_COST = -1 * CARRY_COST
    End If
'---------------------------------------------------------------------------
    If CARRY_COST >= RISK_FREE_RATE Then
    'Never optimal to exersice before maturity
        APPROXIMATION_AMERICAN_OPTION_FUNC = _
            GENERALIZED_BLACK_SCHOLES_FUNC(SPOT, STRIKE, EXPIRATION, _
            RISK_FREE_RATE, CARRY_COST, VOLATILITY, OPTION_FLAG, CND_TYPE)
        Exit Function
    End If
'---------------------------------------------------------------------------

    BETAT_VAL = (1 / 2 - CARRY_COST / VOLATILITY ^ 2) + Sqr((CARRY_COST / _
                VOLATILITY ^ 2 - 1 / 2) ^ 2 + 2 * RISK_FREE_RATE / _
                VOLATILITY ^ 2)
            
    BETAI_VAL = BETAT_VAL / (BETAT_VAL - 1) * STRIKE
    BETA0_VAL = MAXIMUM_FUNC(STRIKE, RISK_FREE_RATE / (RISK_FREE_RATE - _
                CARRY_COST) * STRIKE)

    HT_VAL = -(CARRY_COST * EXPIRATION + 2 * VOLATILITY * Sqr(EXPIRATION)) * _
            BETA0_VAL / (BETAI_VAL - BETA0_VAL)
            
    I_VAL = BETA0_VAL + (BETAI_VAL - BETA0_VAL) * (1 - Exp(HT_VAL))
            
    ALPHAT_VAL = (I_VAL - STRIKE) * I_VAL ^ (-BETAT_VAL)
            
    If SPOT >= I_VAL Then
        APPROXIMATION_AMERICAN_OPTION_FUNC = SPOT - STRIKE
    Else
        TEMP3_VAL = ALPHAT_VAL * SPOT ^ BETAT_VAL - ALPHAT_VAL * _
            CALL_AMERICAN_OPTION_PHI_FUNC(SPOT, EXPIRATION, 0, _
            BETAT_VAL, I_VAL, I_VAL, 0, RISK_FREE_RATE, _
            CARRY_COST, VOLATILITY, 0, CND_TYPE, CBND_TYPE) + _
            CALL_AMERICAN_OPTION_PHI_FUNC(SPOT, EXPIRATION, 0, 1, _
            I_VAL, I_VAL, 0, RISK_FREE_RATE, _
            CARRY_COST, VOLATILITY, 0, CND_TYPE, CBND_TYPE) - _
            CALL_AMERICAN_OPTION_PHI_FUNC(SPOT, EXPIRATION, 0, 1, _
            STRIKE, I_VAL, 0, RISK_FREE_RATE, CARRY_COST, VOLATILITY, _
            0, CND_TYPE, CBND_TYPE) - STRIKE * _
            CALL_AMERICAN_OPTION_PHI_FUNC(SPOT, EXPIRATION, 0, 0, _
            I_VAL, I_VAL, 0, RISK_FREE_RATE, _
            CARRY_COST, VOLATILITY, 0, CND_TYPE, CBND_TYPE) + STRIKE * _
            CALL_AMERICAN_OPTION_PHI_FUNC(SPOT, EXPIRATION, 0, 0, _
            STRIKE, I_VAL, 0, RISK_FREE_RATE, _
            CARRY_COST, VOLATILITY, 0, CND_TYPE, CBND_TYPE)
                      
        APPROXIMATION_AMERICAN_OPTION_FUNC = TEMP3_VAL
    End If

'------------------------------------------------------------------------------
Case 1 'Bjerksund and Stensland American Approximation (2002)
'------------------------------------------------------------------------------

     If OPTION_FLAG = -1 Then
        TEMP1_VAL = SPOT
        TEMP2_VAL = STRIKE
        
        STRIKE = TEMP1_VAL
        SPOT = TEMP2_VAL
        
        RISK_FREE_RATE = RISK_FREE_RATE - CARRY_COST
        CARRY_COST = -1 * CARRY_COST
     End If
     
'---------------------------------------------------------------------------
    If CARRY_COST >= RISK_FREE_RATE Then
    'Never optimal to exersice before maturity
        APPROXIMATION_AMERICAN_OPTION_FUNC = _
            GENERALIZED_BLACK_SCHOLES_FUNC(SPOT, STRIKE, EXPIRATION, _
            RISK_FREE_RATE, CARRY_COST, VOLATILITY, OPTION_FLAG, CND_TYPE)
        Exit Function
    End If
'---------------------------------------------------------------------------

    T_VAL = 1 / 2 * (Sqr(5) - 1) * EXPIRATION

    BETAT_VAL = (1 / 2 - CARRY_COST / VOLATILITY ^ 2) + Sqr((CARRY_COST / _
    VOLATILITY ^ 2 - 1 / 2) ^ 2 + 2 * RISK_FREE_RATE / VOLATILITY ^ 2)
            
    BETAI_VAL = BETAT_VAL / (BETAT_VAL - 1) * STRIKE
    BETA0_VAL = MAXIMUM_FUNC(STRIKE, RISK_FREE_RATE / _
                (RISK_FREE_RATE - CARRY_COST) * STRIKE)

    HT1_VAL = -(CARRY_COST * T_VAL + 2 * VOLATILITY * Sqr(T_VAL)) _
            * STRIKE ^ 2 / ((BETAI_VAL - BETA0_VAL) * BETA0_VAL)
            
    HT2_VAL = -(CARRY_COST * EXPIRATION + 2 * VOLATILITY * Sqr(EXPIRATION)) * _
                STRIKE ^ 2 / ((BETAI_VAL - BETA0_VAL) * BETA0_VAL)
    
    I1_VAL = BETA0_VAL + (BETAI_VAL - BETA0_VAL) * (1 - Exp(HT1_VAL))
    I2_VAL = BETA0_VAL + (BETAI_VAL - BETA0_VAL) * (1 - Exp(HT2_VAL))
    
    ALPHA1_VAL = (I1_VAL - STRIKE) * I1_VAL ^ (-BETAT_VAL)
    ALPHA2_VAL = (I2_VAL - STRIKE) * I2_VAL ^ (-BETAT_VAL)

    If SPOT >= I2_VAL Then
        APPROXIMATION_AMERICAN_OPTION_FUNC = SPOT - STRIKE
    Else
        TEMP3_VAL = (ALPHA2_VAL * SPOT ^ BETAT_VAL) - (ALPHA2_VAL * _
                    CALL_AMERICAN_OPTION_PHI_FUNC(SPOT, T_VAL, 0, _
                    BETAT_VAL, I2_VAL, I2_VAL, 0, RISK_FREE_RATE, _
                    CARRY_COST, VOLATILITY, 1, CND_TYPE, CBND_TYPE))

        TEMP3_VAL = TEMP3_VAL + CALL_AMERICAN_OPTION_PHI_FUNC(SPOT, _
                    T_VAL, 0, 1, I2_VAL, I2_VAL, 0, RISK_FREE_RATE, _
                    CARRY_COST, VOLATILITY, 1, CND_TYPE, CBND_TYPE) - _
                    CALL_AMERICAN_OPTION_PHI_FUNC(SPOT, T_VAL, 0, 1, _
                    I1_VAL, I2_VAL, 0, RISK_FREE_RATE, CARRY_COST, _
                    VOLATILITY, 1, CND_TYPE, CBND_TYPE)
    
        TEMP3_VAL = TEMP3_VAL - STRIKE * _
                    CALL_AMERICAN_OPTION_PHI_FUNC(SPOT, T_VAL, 0, 0, _
                    I2_VAL, I2_VAL, 0, RISK_FREE_RATE, _
                    CARRY_COST, VOLATILITY, 1, CND_TYPE, CBND_TYPE) + _
                    STRIKE * CALL_AMERICAN_OPTION_PHI_FUNC(SPOT, T_VAL, _
                    0, 0, I1_VAL, I2_VAL, 0, RISK_FREE_RATE, _
                    CARRY_COST, VOLATILITY, 1, CND_TYPE, CBND_TYPE)
    
        TEMP3_VAL = TEMP3_VAL + ALPHA1_VAL * _
                    CALL_AMERICAN_OPTION_PHI_FUNC(SPOT, T_VAL, 0, _
                    BETAT_VAL, I1_VAL, I2_VAL, 0, RISK_FREE_RATE, _
                    CARRY_COST, VOLATILITY, 1, CND_TYPE, CBND_TYPE) - _
                    ALPHA1_VAL * CALL_AMERICAN_OPTION_PHI_FUNC(SPOT, _
                    EXPIRATION, T_VAL, BETAT_VAL, I1_VAL, I2_VAL, _
                    I1_VAL, RISK_FREE_RATE, CARRY_COST, VOLATILITY, 2, _
                    CND_TYPE, CBND_TYPE)
    
        TEMP3_VAL = TEMP3_VAL + _
                    CALL_AMERICAN_OPTION_PHI_FUNC(SPOT, EXPIRATION, _
                    T_VAL, 1, I1_VAL, I2_VAL, I1_VAL, RISK_FREE_RATE, _
                    CARRY_COST, VOLATILITY, 2, CND_TYPE, CBND_TYPE) - _
                    CALL_AMERICAN_OPTION_PHI_FUNC(SPOT, EXPIRATION, _
                    T_VAL, 1, STRIKE, I2_VAL, I1_VAL, RISK_FREE_RATE, _
                    CARRY_COST, VOLATILITY, 2, CND_TYPE, CBND_TYPE)
                   
        TEMP3_VAL = TEMP3_VAL - STRIKE * _
                    CALL_AMERICAN_OPTION_PHI_FUNC(SPOT, EXPIRATION, T_VAL, _
                    0, I1_VAL, I2_VAL, I1_VAL, RISK_FREE_RATE, CARRY_COST, _
                    VOLATILITY, 2, CND_TYPE, CBND_TYPE) + STRIKE * _
                    CALL_AMERICAN_OPTION_PHI_FUNC(SPOT, EXPIRATION, T_VAL, 0, _
                    STRIKE, I2_VAL, I1_VAL, RISK_FREE_RATE, CARRY_COST, _
                    VOLATILITY, 2, CND_TYPE, CBND_TYPE)
    
        APPROXIMATION_AMERICAN_OPTION_FUNC = TEMP3_VAL
    End If

'------------------------------------------------------------------------------
Case Else 'Barone-Adesi and Whaley (1987) American approximation
'------------------------------------------------------------------------------
    Select Case OPTION_FLAG
'---------------------------------------------------------------------------
    Case 1 ', "CALL", "C"
'---------------------------------------------------------------------------
        If CARRY_COST >= RISK_FREE_RATE Then
        'Never optimal to exersice before maturity
            APPROXIMATION_AMERICAN_OPTION_FUNC = _
            GENERALIZED_BLACK_SCHOLES_FUNC(SPOT, STRIKE, EXPIRATION, _
            RISK_FREE_RATE, CARRY_COST, VOLATILITY, 1, CND_TYPE)
            Exit Function
        End If
'---------------------------------------------------------------------------
        SK_VAL = CALL_AMERICAN_OPTION_NEWTON_FUNC(STRIKE, EXPIRATION, _
                 RISK_FREE_RATE, CARRY_COST, VOLATILITY, 1, CND_TYPE)
             
        N_VAL = 2 * CARRY_COST / VOLATILITY ^ 2
        K_VAL = 2 * RISK_FREE_RATE / (VOLATILITY ^ 2 * (1 - _
                Exp(-RISK_FREE_RATE * EXPIRATION)))
             
        D_VAL = (Log(SK_VAL / STRIKE) + (CARRY_COST + VOLATILITY ^ 2 / 2) * _
                EXPIRATION) / (VOLATILITY * (EXPIRATION) ^ 0.5)
        
        Q_VAL = (-(N_VAL - 1) + Sqr((N_VAL - 1) ^ 2 + 4 * K_VAL)) / 2
        A_VAL = (SK_VAL / Q_VAL) * (1 - Exp((CARRY_COST - _
                RISK_FREE_RATE) * EXPIRATION) * CND_FUNC(D_VAL, CND_TYPE))
             
        If SPOT < SK_VAL Then
            APPROXIMATION_AMERICAN_OPTION_FUNC = _
                    GENERALIZED_BLACK_SCHOLES_FUNC(SPOT, STRIKE, _
                    EXPIRATION, RISK_FREE_RATE, CARRY_COST, VOLATILITY, _
                    1, CND_TYPE) + A_VAL * (SPOT / SK_VAL) ^ Q_VAL
        Else
            APPROXIMATION_AMERICAN_OPTION_FUNC = SPOT - STRIKE
        End If
'---------------------------------------------------------------------------
    Case Else '-1, "PUT", "P"
'---------------------------------------------------------------------------
        SK_VAL = CALL_AMERICAN_OPTION_NEWTON_FUNC(STRIKE, EXPIRATION, _
                 RISK_FREE_RATE, CARRY_COST, VOLATILITY, -1, CND_TYPE)
        
        N_VAL = 2 * CARRY_COST / VOLATILITY ^ 2
        K_VAL = 2 * RISK_FREE_RATE / (VOLATILITY ^ 2 * (1 - _
                Exp(-RISK_FREE_RATE * EXPIRATION)))
        D_VAL = (Log(SK_VAL / STRIKE) + (CARRY_COST + VOLATILITY ^ 2 / 2) * _
                EXPIRATION) / (VOLATILITY * (EXPIRATION) ^ 0.5)
        Q_VAL = (-(N_VAL - 1) - Sqr((N_VAL - 1) ^ 2 + 4 * K_VAL)) / 2
        A_VAL = -(SK_VAL / Q_VAL) * (1 - Exp((CARRY_COST - RISK_FREE_RATE) * _
                EXPIRATION) * CND_FUNC(-D_VAL, CND_TYPE))
        If SPOT > SK_VAL Then
            APPROXIMATION_AMERICAN_OPTION_FUNC = _
                GENERALIZED_BLACK_SCHOLES_FUNC(SPOT, STRIKE, _
                EXPIRATION, RISK_FREE_RATE, CARRY_COST, VOLATILITY, -1, _
                CND_TYPE) + A_VAL * (SPOT / SK_VAL) ^ Q_VAL
        Else
            APPROXIMATION_AMERICAN_OPTION_FUNC = STRIKE - SPOT
        End If
'---------------------------------------------------------------------------
    End Select
'---------------------------------------------------------------------------
End Select
'---------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
APPROXIMATION_AMERICAN_OPTION_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : ROLL_GESKE_WHALEY_AMERICAN_OPTION_FUNC
'DESCRIPTION   : American Calls on stocks with known dividends, Roll-Geske-Whaley
'LIBRARY       : DERIVATIVES
'GROUP         : AMER_APPROX
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function ROLL_GESKE_WHALEY_AMERICAN_OPTION_FUNC(ByVal SPOT As Double, _
ByVal STRIKE As Double, _
ByVal DIVD_TENOR As Double, _
ByVal EXPIRATION As Double, _
ByVal RISK_FREE_RATE As Double, _
ByVal CASH_DIVIDEND_SHARE As Double, _
ByVal VOLATILITY As Double, _
Optional ByVal CND_TYPE As Integer = 0, _
Optional ByVal CBND_TYPE As Integer = 0)
    
'DIVD_TENOR time to dividend payout
'EXPIRATION time to option expiration
'CASH_DIVIDEND_SHARE : Dollar Amount
    
Dim I_VAL As Double
Dim A1_VAL As Double
Dim A2_VAL As Double

Dim B1_VAL As Double
Dim B2_VAL As Double
Dim BS_VAL As Double

Dim HIGH_VAL As Double
Dim LOW_VAL As Double
Dim TEMP_VAL As Double

Dim LOWER_BOUND As Double
Dim UPPER_BOUND As Double

On Error GoTo ERROR_LABEL

UPPER_BOUND = 100000000
LOWER_BOUND = 0.00001

TEMP_VAL = SPOT - CASH_DIVIDEND_SHARE * Exp(-RISK_FREE_RATE * DIVD_TENOR)

If CASH_DIVIDEND_SHARE <= STRIKE * (1 - Exp(-RISK_FREE_RATE * _
    (EXPIRATION - DIVD_TENOR))) Then _
'-------------------> Not optimal to exercise
    ROLL_GESKE_WHALEY_AMERICAN_OPTION_FUNC = _
        BLACK_SCHOLES_OPTION_FUNC(TEMP_VAL, STRIKE, EXPIRATION, _
        RISK_FREE_RATE, VOLATILITY, 1)
    Exit Function
End If

BS_VAL = BLACK_SCHOLES_OPTION_FUNC(SPOT, STRIKE, EXPIRATION - DIVD_TENOR, _
         RISK_FREE_RATE, VOLATILITY, 1)
HIGH_VAL = SPOT

Do While (BS_VAL - HIGH_VAL - CASH_DIVIDEND_SHARE + STRIKE) > 0 And _
    HIGH_VAL < UPPER_BOUND
    HIGH_VAL = HIGH_VAL * 2
    BS_VAL = BLACK_SCHOLES_OPTION_FUNC(HIGH_VAL, STRIKE, _
             EXPIRATION - DIVD_TENOR, _
    RISK_FREE_RATE, VOLATILITY, 1)
Loop
If HIGH_VAL > UPPER_BOUND Then
    ROLL_GESKE_WHALEY_AMERICAN_OPTION_FUNC = _
        BLACK_SCHOLES_OPTION_FUNC(TEMP_VAL, STRIKE, EXPIRATION, _
        RISK_FREE_RATE, VOLATILITY, 1)
    Exit Function
End If

LOW_VAL = 0
I_VAL = HIGH_VAL * 0.5
BS_VAL = BLACK_SCHOLES_OPTION_FUNC(I_VAL, STRIKE, EXPIRATION - _
         DIVD_TENOR, RISK_FREE_RATE, VOLATILITY, 1)

Do While Abs(BS_VAL - I_VAL - CASH_DIVIDEND_SHARE + STRIKE) > LOWER_BOUND _
    And HIGH_VAL - LOW_VAL > LOWER_BOUND
    'search newton algorithm to find the critical asset price
    If (BS_VAL - I_VAL - CASH_DIVIDEND_SHARE + STRIKE) < 0 Then
        HIGH_VAL = I_VAL
    Else
        LOW_VAL = I_VAL
    End If
    I_VAL = (HIGH_VAL + LOW_VAL) / 2
    BS_VAL = BLACK_SCHOLES_OPTION_FUNC(I_VAL, STRIKE, EXPIRATION - _
             DIVD_TENOR, RISK_FREE_RATE, VOLATILITY, 1)
Loop

A1_VAL = (Log(TEMP_VAL / STRIKE) + (RISK_FREE_RATE + _
         VOLATILITY ^ 2 / 2) * EXPIRATION) / _
        (VOLATILITY * Sqr(EXPIRATION))
A2_VAL = A1_VAL - VOLATILITY * Sqr(EXPIRATION)
B1_VAL = (Log(TEMP_VAL / I_VAL) + (RISK_FREE_RATE + _
        VOLATILITY ^ 2 / 2) * DIVD_TENOR) / _
       (VOLATILITY * Sqr(DIVD_TENOR))
B2_VAL = B1_VAL - VOLATILITY * Sqr(DIVD_TENOR)

ROLL_GESKE_WHALEY_AMERICAN_OPTION_FUNC = _
        TEMP_VAL * CND_FUNC(B1_VAL, CND_TYPE) + TEMP_VAL * _
        CBND_FUNC(A1_VAL, -B1_VAL, -Sqr(DIVD_TENOR / EXPIRATION), _
        CND_TYPE, CBND_TYPE) - STRIKE * Exp(-RISK_FREE_RATE * _
        EXPIRATION) * CBND_FUNC(A2_VAL, -B2_VAL, _
        -Sqr(DIVD_TENOR / EXPIRATION), CND_TYPE, CBND_TYPE) - _
        (STRIKE - CASH_DIVIDEND_SHARE) * Exp(-RISK_FREE_RATE * _
        DIVD_TENOR) * CND_FUNC(B2_VAL, CND_TYPE)

Exit Function
ERROR_LABEL:
ROLL_GESKE_WHALEY_AMERICAN_OPTION_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_AMERICAN_OPTION_NEWTON_FUNC
'DESCRIPTION   : Newton Raphson algorithm to solve for the critical American Price
'LIBRARY       : DERIVATIVES
'GROUP         : AMER_APPROX
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Private Function CALL_AMERICAN_OPTION_NEWTON_FUNC(ByVal STRIKE As Double, _
ByVal EXPIRATION As Double, _
ByVal RISK_FREE_RATE As Double, _
ByVal CARRY_COST As Double, _
ByVal VOLATILITY As Double, _
Optional ByVal OPTION_FLAG As Integer = 1, _
Optional ByVal CND_TYPE As Integer = 0)

Dim i As Long
Dim j As Long
    
Dim M_VAL As Double
Dim N_VAL As Double
    
Dim H_VAL As Double
Dim K_VAL As Double
    
Dim D_VAL As Double
Dim Q_VAL As Double
Dim U_VAL As Double
        
Dim SU_VAL As Double
Dim SI_VAL As Double
        
Dim LHS_VAL As Double
Dim RHS_VAL As Double
Dim BI_VAL As Double
        
Dim tolerance As Double
    
On Error GoTo ERROR_LABEL
    
tolerance = 0.000001 '---> You can change this root for the Newton _
search Algorithm
    
j = 100 'limit for the loop
    
'--------------------------------------------------------------------------------
Select Case OPTION_FLAG
'--------------------------------------------------------------------------------
Case 1 ', "CALL", "C"
'--------------------------------------------------------------------------------
    N_VAL = 2 * CARRY_COST / VOLATILITY ^ 2
    M_VAL = 2 * RISK_FREE_RATE / VOLATILITY ^ 2
    U_VAL = (-(N_VAL - 1) + Sqr((N_VAL - 1) ^ 2 + 4 * M_VAL)) / 2
    
    SU_VAL = STRIKE / (1 - 1 / U_VAL)
    H_VAL = -(CARRY_COST * EXPIRATION + 2 * VOLATILITY * Sqr(EXPIRATION)) * _
        STRIKE / (SU_VAL - STRIKE)
    
    SI_VAL = STRIKE + (SU_VAL - STRIKE) * (1 - Exp(H_VAL))

    K_VAL = 2 * RISK_FREE_RATE / (VOLATILITY ^ 2 * (1 - _
            Exp(-RISK_FREE_RATE * EXPIRATION)))
    
    D_VAL = (Log(SI_VAL / STRIKE) + (CARRY_COST + VOLATILITY ^ 2 / 2) * _
        EXPIRATION) / (VOLATILITY * Sqr(EXPIRATION))
    
    Q_VAL = (-(N_VAL - 1) + Sqr((N_VAL - 1) ^ 2 + 4 * K_VAL)) / 2
    LHS_VAL = SI_VAL - STRIKE
    
    RHS_VAL = GENERALIZED_BLACK_SCHOLES_FUNC(SI_VAL, STRIKE, _
              EXPIRATION, RISK_FREE_RATE, CARRY_COST, _
              VOLATILITY, 1, CND_TYPE) + (1 - Exp((CARRY_COST - _
              RISK_FREE_RATE) * EXPIRATION) * CND_FUNC(D_VAL, _
              CND_TYPE)) * SI_VAL / Q_VAL
    
    BI_VAL = Exp((CARRY_COST - RISK_FREE_RATE) * EXPIRATION) * _
             CND_FUNC(D_VAL, CND_TYPE) * (1 - 1 / Q_VAL) + (1 - _
             Exp((CARRY_COST - RISK_FREE_RATE) * EXPIRATION) * _
             CND_FUNC(D_VAL, CND_TYPE) / (VOLATILITY * Sqr(EXPIRATION))) / Q_VAL
    
    i = 0
    
    Do While Abs(LHS_VAL - RHS_VAL) / STRIKE > tolerance
    'Newton Raphson algorithm for finding critical price.
        
        SI_VAL = (STRIKE + RHS_VAL - BI_VAL * SI_VAL) / (1 - BI_VAL)
        D_VAL = (Log(SI_VAL / STRIKE) + (CARRY_COST + VOLATILITY ^ 2 / 2) * _
                EXPIRATION) / (VOLATILITY * Sqr(EXPIRATION))
        
        LHS_VAL = SI_VAL - STRIKE
        RHS_VAL = GENERALIZED_BLACK_SCHOLES_FUNC(SI_VAL, STRIKE, _
                  EXPIRATION, RISK_FREE_RATE, CARRY_COST, _
                  VOLATILITY, 1, CND_TYPE) + (1 - Exp((CARRY_COST - _
                  RISK_FREE_RATE) * EXPIRATION) * CND_FUNC(D_VAL, _
                  CND_TYPE)) * SI_VAL / Q_VAL
        
        BI_VAL = Exp((CARRY_COST - RISK_FREE_RATE) * EXPIRATION) * _
                 CND_FUNC(D_VAL, CND_TYPE) * (1 - 1 / Q_VAL) + (1 - _
                 Exp((CARRY_COST - RISK_FREE_RATE) * EXPIRATION) * _
                NORMAL_MASS_DIST_FUNC(D_VAL) / (VOLATILITY * _
                Sqr(EXPIRATION))) / Q_VAL
        i = i + 1
        
        If i > j Then: GoTo ERROR_LABEL
    
    Loop
    
    CALL_AMERICAN_OPTION_NEWTON_FUNC = SI_VAL
    
'--------------------------------------------------------------------------------
Case Else '-1, "PUT", "P"
'--------------------------------------------------------------------------------

    N_VAL = 2 * CARRY_COST / VOLATILITY ^ 2
    M_VAL = 2 * RISK_FREE_RATE / VOLATILITY ^ 2
    U_VAL = (-(N_VAL - 1) - Sqr((N_VAL - 1) ^ 2 + 4 * M_VAL)) / 2
    SU_VAL = STRIKE / (1 - 1 / U_VAL)
    
    H_VAL = (CARRY_COST * EXPIRATION - 2 * VOLATILITY * _
            Sqr(EXPIRATION)) * STRIKE / (STRIKE - SU_VAL)
    
    SI_VAL = SU_VAL + (STRIKE - SU_VAL) * Exp(H_VAL)
    
    K_VAL = 2 * RISK_FREE_RATE / (VOLATILITY ^ 2 * (1 - _
            Exp(-RISK_FREE_RATE * EXPIRATION)))
    
    D_VAL = (Log(SI_VAL / STRIKE) + (CARRY_COST + VOLATILITY ^ 2 / 2) * _
            EXPIRATION) / (VOLATILITY * Sqr(EXPIRATION))
    
    Q_VAL = (-(N_VAL - 1) - Sqr((N_VAL - 1) ^ 2 + 4 * K_VAL)) / 2
    
    LHS_VAL = STRIKE - SI_VAL
    RHS_VAL = GENERALIZED_BLACK_SCHOLES_FUNC(SI_VAL, STRIKE, EXPIRATION, _
              RISK_FREE_RATE, CARRY_COST, VOLATILITY, -1, CND_TYPE) - _
              (1 - Exp((CARRY_COST - RISK_FREE_RATE) * EXPIRATION) * _
              CND_FUNC(-D_VAL, CND_TYPE)) * SI_VAL / Q_VAL
    
    BI_VAL = -Exp((CARRY_COST - RISK_FREE_RATE) * EXPIRATION) * _
            CND_FUNC(-D_VAL, CND_TYPE) * (1 - 1 / Q_VAL) - (1 + _
            Exp((CARRY_COST - RISK_FREE_RATE) * EXPIRATION) * _
            NORMAL_MASS_DIST_FUNC(-D_VAL) / _
            (VOLATILITY * Sqr(EXPIRATION))) / Q_VAL

    i = 0
    Do While Abs(LHS_VAL - RHS_VAL) / STRIKE > tolerance
    'Newton Raphson algorithm for finding critical price.
        SI_VAL = (STRIKE - RHS_VAL + BI_VAL * SI_VAL) / (1 + BI_VAL)
        
        D_VAL = (Log(SI_VAL / STRIKE) + (CARRY_COST + VOLATILITY ^ 2 / 2) * _
            EXPIRATION) / (VOLATILITY * Sqr(EXPIRATION))
        
        LHS_VAL = STRIKE - SI_VAL
        
        RHS_VAL = GENERALIZED_BLACK_SCHOLES_FUNC(SI_VAL, STRIKE, _
                  EXPIRATION, RISK_FREE_RATE, CARRY_COST, VOLATILITY, _
                  -1, CND_TYPE) - (1 - Exp((CARRY_COST - RISK_FREE_RATE) * _
                  EXPIRATION) * CND_FUNC(-D_VAL, CND_TYPE)) * SI_VAL / Q_VAL
        
        BI_VAL = -Exp((CARRY_COST - RISK_FREE_RATE) * EXPIRATION) * _
                 CND_FUNC(-D_VAL, CND_TYPE) * (1 - 1 / Q_VAL) _
                 - (1 + Exp((CARRY_COST - RISK_FREE_RATE) * EXPIRATION) * _
                 CND_FUNC(-D_VAL, CND_TYPE) / (VOLATILITY * _
                 Sqr(EXPIRATION))) / Q_VAL

        i = i + 1
        
        If i > j Then: GoTo ERROR_LABEL

    Loop

CALL_AMERICAN_OPTION_NEWTON_FUNC = SI_VAL
    
'--------------------------------------------------------------------------------
End Select
'--------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
CALL_AMERICAN_OPTION_NEWTON_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_AMERICAN_OPTION_PHI_FUNC
'DESCRIPTION   : American Phi Factor
'LIBRARY       : DERIVATIVES
'GROUP         : AMER_APPROX
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Private Function CALL_AMERICAN_OPTION_PHI_FUNC(ByVal S_VAL As Double, _
ByVal T2_VAL As Double, _
ByVal T1_VAL As Double, _
ByVal GAMMA_VAL As Double, _
ByVal H_VAL As Double, _
ByVal I2_VAL As Double, _
ByVal I1_VAL As Double, _
ByVal R_VAL As Double, _
ByVal B_VAL As Double, _
ByVal V_VAL As Double, _
Optional ByVal VERSION As Integer = 0, _
Optional ByVal CND_TYPE As Integer = 0, _
Optional ByVal CBND_TYPE As Integer = 0)

Dim D1_VAL As Double
Dim D2_VAL As Double

Dim E1_VAL As Double
Dim E2_VAL As Double
Dim E3_VAL As Double
Dim E4_VAL As Double
    
Dim F1_VAL As Double
Dim F2_VAL As Double
Dim F3_VAL As Double
Dim F4_VAL As Double
    
Dim RHO_VAL As Double
Dim KAPPA_VAL As Double
Dim LAMBDA_VAL As Double
    
On Error GoTo ERROR_LABEL
    
'-----------------------------------------------------------------------------
Select Case VERSION
'-----------------------------------------------------------------------------
Case 0
'-----------------------------------------------------------------------------
    LAMBDA_VAL = (-R_VAL + GAMMA_VAL * B_VAL + 0.5 * _
                GAMMA_VAL * (GAMMA_VAL - 1) * V_VAL ^ 2) * T2_VAL
    
    D1_VAL = -(Log(S_VAL / H_VAL) + (B_VAL + (GAMMA_VAL - 0.5) * _
            V_VAL ^ 2) * T2_VAL) / (V_VAL * Sqr(T2_VAL))
    
    KAPPA_VAL = 2 * B_VAL / (V_VAL ^ 2) + (2 * GAMMA_VAL - 1)
    
    CALL_AMERICAN_OPTION_PHI_FUNC = Exp(LAMBDA_VAL) * S_VAL ^ _
        GAMMA_VAL * (CND_FUNC(D1_VAL, CND_TYPE) - _
        (I2_VAL / S_VAL) ^ KAPPA_VAL * CND_FUNC(D1_VAL - 2 * _
        Log(I2_VAL / S_VAL) / _
        (V_VAL * Sqr(T2_VAL)), CND_TYPE))
'-----------------------------------------------------------------------------
Case 1
'-----------------------------------------------------------------------------
    LAMBDA_VAL = -R_VAL + GAMMA_VAL * B_VAL + 0.5 * GAMMA_VAL * _
                (GAMMA_VAL - 1) * V_VAL ^ 2
    KAPPA_VAL = 2 * B_VAL / (V_VAL ^ 2) + (2 * GAMMA_VAL - 1)
    
    D1_VAL = (Log(S_VAL / H_VAL) + (B_VAL + (GAMMA_VAL - 0.5) * _
            V_VAL ^ 2) * T2_VAL) / (V_VAL * Sqr(T2_VAL))
    
    D2_VAL = (Log(I2_VAL ^ 2 / (S_VAL * H_VAL)) + (B_VAL + _
            (GAMMA_VAL - 0.5) * V_VAL ^ 2) * T2_VAL) / (V_VAL _
            * Sqr(T2_VAL))
    
    CALL_AMERICAN_OPTION_PHI_FUNC = _
        Exp(LAMBDA_VAL * T2_VAL) * S_VAL ^ GAMMA_VAL * (CND_FUNC(-D1_VAL, _
        CND_TYPE) - (I2_VAL / S_VAL) ^ KAPPA_VAL * CND_FUNC(-D2_VAL, CND_TYPE))
'-----------------------------------------------------------------------------
Case Else '2
'-----------------------------------------------------------------------------
    E1_VAL = (Log(S_VAL / I1_VAL) + (B_VAL + (GAMMA_VAL - 0.5) * _
              V_VAL ^ 2) * T1_VAL) / (V_VAL * Sqr(T1_VAL))
    
    E2_VAL = (Log(I2_VAL ^ 2 / (S_VAL * I1_VAL)) + (B_VAL + _
             (GAMMA_VAL - 0.5) * V_VAL ^ 2) * T1_VAL) / _
             (V_VAL * Sqr(T1_VAL))
 
    E3_VAL = (Log(S_VAL / I1_VAL) - (B_VAL + (GAMMA_VAL - 0.5) * _
             V_VAL ^ 2) * T1_VAL) / (V_VAL * Sqr(T1_VAL))
        
    E4_VAL = (Log(I2_VAL ^ 2 / (S_VAL * I1_VAL)) - (B_VAL + _
             (GAMMA_VAL - 0.5) * V_VAL ^ 2) * T1_VAL) / _
             (V_VAL * Sqr(T1_VAL))

    F1_VAL = (Log(S_VAL / H_VAL) + (B_VAL + (GAMMA_VAL - 0.5) * _
              V_VAL ^ 2) * T2_VAL) / (V_VAL * Sqr(T2_VAL))
    
    F2_VAL = (Log(I2_VAL ^ 2 / (S_VAL * H_VAL)) + (B_VAL + _
             (GAMMA_VAL - 0.5) * V_VAL ^ 2) * T2_VAL) / _
             (V_VAL * Sqr(T2_VAL))
    
    F3_VAL = (Log(I1_VAL ^ 2 / (S_VAL * H_VAL)) + (B_VAL + _
             (GAMMA_VAL - 0.5) * V_VAL ^ 2) * T2_VAL) / _
             (V_VAL * Sqr(T2_VAL))
   
    F4_VAL = (Log((S_VAL * I1_VAL ^ 2) / (H_VAL * I2_VAL ^ 2)) + _
             (B_VAL + (GAMMA_VAL - 0.5) * V_VAL ^ 2) _
             * T2_VAL) / (V_VAL * Sqr(T2_VAL))
       
    RHO_VAL = Sqr(T1_VAL / T2_VAL)
    
    LAMBDA_VAL = -R_VAL + GAMMA_VAL * B_VAL + 0.5 * GAMMA_VAL * _
                 (GAMMA_VAL - 1) * V_VAL ^ 2
    
    KAPPA_VAL = 2 * B_VAL / (V_VAL ^ 2) + (2 * GAMMA_VAL - 1)

    CALL_AMERICAN_OPTION_PHI_FUNC = _
        Exp(LAMBDA_VAL * T2_VAL) * S_VAL ^ GAMMA_VAL * (CBND_FUNC(-E1_VAL, _
        -F1_VAL, RHO_VAL, CND_TYPE, CBND_TYPE) - (I2_VAL / S_VAL) ^ _
        KAPPA_VAL * CBND_FUNC(-E2_VAL, -F2_VAL, RHO_VAL, CND_TYPE, _
        CBND_TYPE) - (I1_VAL / S_VAL) ^ KAPPA_VAL * CBND_FUNC(-E3_VAL, _
        -F3_VAL, -RHO_VAL, CND_TYPE, CBND_TYPE) + (I1_VAL / I2_VAL) ^ _
        KAPPA_VAL * CBND_FUNC(-E4_VAL, -F4_VAL, -RHO_VAL, CND_TYPE, CBND_TYPE))
'-----------------------------------------------------------------------------
End Select
'-----------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
CALL_AMERICAN_OPTION_PHI_FUNC = Err.number
End Function
