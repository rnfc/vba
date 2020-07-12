Attribute VB_Name = "FINAN_DERIV_CHOOSER_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'**********************************************************************************
'**********************************************************************************
'FUNCTION      : SIMPLE_CHOOSER_OPTION_FUNC
'DESCRIPTION   : SIMPLE_CHOOSER_OPT
'LIBRARY       : DERIVATIVES
'GROUP         : CHOOSER
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'**********************************************************************************
'**********************************************************************************

Function SIMPLE_CHOOSER_OPTION_FUNC(ByVal SPOT As Double, _
ByVal STRIKE As Double, _
ByVal CHOOSER_TENOR As Double, _
ByVal EXPIRATION As Double, _
ByVal RATE As Double, _
ByVal CARRY_COST As Double, _
ByVal SIGMA As Double, _
Optional ByVal CND_TYPE As Integer = 0)

    Dim ATEMP_VAL As Double
    Dim BTEMP_VAL As Double

    On Error GoTo ERROR_LABEL
    
    ATEMP_VAL = (Log(SPOT / STRIKE) + (CARRY_COST + _
         SIGMA ^ 2 / 2) * EXPIRATION) / _
        (SIGMA * Sqr(EXPIRATION))
    
    BTEMP_VAL = (Log(SPOT / STRIKE) + CARRY_COST * EXPIRATION + SIGMA ^ 2 * _
        CHOOSER_TENOR / 2) / (SIGMA * Sqr(CHOOSER_TENOR))
  
    SIMPLE_CHOOSER_OPTION_FUNC = SPOT * Exp((CARRY_COST - RATE) * EXPIRATION) * _
        CND_FUNC(ATEMP_VAL, CND_TYPE) - STRIKE * Exp(-RATE * EXPIRATION) * _
        CND_FUNC(ATEMP_VAL - SIGMA * Sqr(EXPIRATION), CND_TYPE) - SPOT * _
        Exp((CARRY_COST - RATE) * EXPIRATION) * _
        CND_FUNC(-BTEMP_VAL, CND_TYPE) + STRIKE * Exp(-RATE * EXPIRATION) * _
        CND_FUNC(-BTEMP_VAL + SIGMA * Sqr(CHOOSER_TENOR), CND_TYPE)
    
Exit Function
ERROR_LABEL:
SIMPLE_CHOOSER_OPTION_FUNC = Err.number
End Function


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : COMPLEX_CHOOSER_OPTION_FUNC
'DESCRIPTION   : COMPLEX_CHOOSER_OPT
'LIBRARY       : DERIVATIVES
'GROUP         : CHOOSER
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'**********************************************************************************
'**********************************************************************************

Function COMPLEX_CHOOSER_OPTION_FUNC(ByVal SPOT As Double, _
ByVal STRIKE_CALL As Double, _
ByVal STRIKE_PUT As Double, _
ByVal CHOOSER_TENOR As Double, _
ByVal CALL_EXPIRATION As Double, _
ByVal PUT_EXPIRATION As Double, _
ByVal RATE As Double, _
ByVal CARRY_COST As Double, _
ByVal SIGMA As Double, _
Optional ByVal CND_TYPE As Integer = 0, _
Optional ByVal CBND_TYPE As Integer = 0)
    
    Dim D1_VAL As Double
    Dim D2_VAL As Double
    
    Dim Y1_VAL As Double
    Dim Y2_VAL As Double
    
    Dim RHO1_VAL As Double
    Dim RHO2_VAL As Double

    Dim DELTA_SPOT_VAL As Double
    
    Dim CALL_VAL As Double
    Dim PUT_VAL As Double
    
    Dim DELTA_CALL_VAL As Double
    Dim DELTA_PUT_VAL As Double
    
    Dim ATEMP_VAL As Double
    Dim BTEMP_VAL As Double
    Dim CRITICAL_VAL As Double

    Dim tolerance As Double
    
    On Error GoTo ERROR_LABEL
    
    DELTA_SPOT_VAL = SPOT
    
    CALL_VAL = GENERALIZED_BLACK_SCHOLES_FUNC(DELTA_SPOT_VAL, STRIKE_CALL, _
        CALL_EXPIRATION - CHOOSER_TENOR, RATE, _
        CARRY_COST, SIGMA, 1, CND_TYPE)
    
    PUT_VAL = GENERALIZED_BLACK_SCHOLES_FUNC(DELTA_SPOT_VAL, STRIKE_PUT, _
        PUT_EXPIRATION - CHOOSER_TENOR, RATE, _
        CARRY_COST, SIGMA, -1, CND_TYPE)
    
        DELTA_CALL_VAL = Exp((CARRY_COST - RATE) * (CALL_EXPIRATION - _
            CHOOSER_TENOR)) * CND_FUNC((Log(DELTA_SPOT_VAL / STRIKE_CALL) + _
            (CARRY_COST + SIGMA ^ 2 / 2) * (CALL_EXPIRATION - _
            CHOOSER_TENOR)) / (SIGMA * Sqr((CALL_EXPIRATION - _
            CHOOSER_TENOR))), CND_TYPE)
    
        DELTA_PUT_VAL = Exp((CARRY_COST - RATE) * (PUT_EXPIRATION - _
            CHOOSER_TENOR)) * (CND_FUNC((Log(DELTA_SPOT_VAL / STRIKE_PUT) + _
            (CARRY_COST + SIGMA ^ 2 / 2) * (PUT_EXPIRATION - _
            CHOOSER_TENOR)) / _
            (SIGMA * Sqr((PUT_EXPIRATION - CHOOSER_TENOR))), _
            CND_TYPE) - 1)
    
    ATEMP_VAL = CALL_VAL - PUT_VAL
    BTEMP_VAL = DELTA_CALL_VAL - DELTA_PUT_VAL
    
    tolerance = 0.001
    
    Do While Abs(ATEMP_VAL) > tolerance 'Newton-Raphson Critical value _
    complex chooser option

        
        DELTA_SPOT_VAL = DELTA_SPOT_VAL - (ATEMP_VAL) / BTEMP_VAL
        
        CALL_VAL = GENERALIZED_BLACK_SCHOLES_FUNC(DELTA_SPOT_VAL, STRIKE_CALL, _
            CALL_EXPIRATION - CHOOSER_TENOR, RATE, _
            CARRY_COST, SIGMA, 1, CND_TYPE)
        
        PUT_VAL = GENERALIZED_BLACK_SCHOLES_FUNC(DELTA_SPOT_VAL, STRIKE_PUT, _
            PUT_EXPIRATION - CHOOSER_TENOR, RATE, _
            CARRY_COST, SIGMA, -1, CND_TYPE)
        
        DELTA_CALL_VAL = Exp((CARRY_COST - RATE) * (CALL_EXPIRATION - _
            CHOOSER_TENOR)) * CND_FUNC((Log(DELTA_SPOT_VAL / STRIKE_CALL) + _
            (CARRY_COST + SIGMA ^ 2 / 2) * (CALL_EXPIRATION - _
            CHOOSER_TENOR)) / (SIGMA * Sqr((CALL_EXPIRATION - _
            CHOOSER_TENOR))), CND_TYPE)
        
        DELTA_PUT_VAL = Exp((CARRY_COST - RATE) * (PUT_EXPIRATION - _
            CHOOSER_TENOR)) * (CND_FUNC((Log(DELTA_SPOT_VAL / STRIKE_PUT) + _
            (CARRY_COST + SIGMA ^ 2 / 2) * (PUT_EXPIRATION - _
            CHOOSER_TENOR)) / (SIGMA * Sqr((PUT_EXPIRATION - _
            CHOOSER_TENOR))), CND_TYPE) - 1)
        
        ATEMP_VAL = CALL_VAL - PUT_VAL
        BTEMP_VAL = DELTA_CALL_VAL - DELTA_PUT_VAL
    
    Loop

    CRITICAL_VAL = DELTA_SPOT_VAL
    
    D1_VAL = (Log(SPOT / CRITICAL_VAL) + (CARRY_COST + SIGMA ^ 2 / 2) * _
      CHOOSER_TENOR) / (SIGMA * Sqr(CHOOSER_TENOR))
    
    D2_VAL = D1_VAL - SIGMA * Sqr(CHOOSER_TENOR)
    Y1_VAL = (Log(SPOT / STRIKE_CALL) + (CARRY_COST + SIGMA ^ 2 / 2) * _
        CALL_EXPIRATION) / (SIGMA * Sqr(CALL_EXPIRATION))
    
    Y2_VAL = (Log(SPOT / STRIKE_PUT) + (CARRY_COST + SIGMA ^ 2 / 2) * _
        PUT_EXPIRATION) / (SIGMA * Sqr(PUT_EXPIRATION))
    
    RHO1_VAL = Sqr(CHOOSER_TENOR / CALL_EXPIRATION)
    RHO2_VAL = Sqr(CHOOSER_TENOR / PUT_EXPIRATION)
    
    COMPLEX_CHOOSER_OPTION_FUNC = SPOT * Exp((CARRY_COST - RATE) * _
        CALL_EXPIRATION) * _
        CBND_FUNC(D1_VAL, Y1_VAL, RHO1_VAL, CND_TYPE, CBND_TYPE) - STRIKE_CALL * _
        Exp(-RATE * CALL_EXPIRATION) * _
        CBND_FUNC(D2_VAL, Y1_VAL - SIGMA * Sqr(CALL_EXPIRATION), RHO1_VAL, _
        CBND_TYPE, CND_TYPE) - SPOT * _
        Exp((CARRY_COST - RATE) * PUT_EXPIRATION) * _
        CBND_FUNC(-D1_VAL, -Y2_VAL, RHO2_VAL, CND_TYPE, CBND_TYPE) + _
        STRIKE_PUT * Exp(-RATE * PUT_EXPIRATION) * _
        CBND_FUNC(-D2_VAL, -Y2_VAL + SIGMA * _
        Sqr(PUT_EXPIRATION), RHO2_VAL, CND_TYPE, CBND_TYPE)

Exit Function
ERROR_LABEL:
COMPLEX_CHOOSER_OPTION_FUNC = Err.number
End Function

'////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////

'About Chooser Options

'Description
'The unique feature of a chooser option is the ability to purchase the
'option now, but not decide until later whether the option is a put
'or a call.

'Types of chooser options:
'1) Complex chooser: Differing tenor and/or strike price for
'the call and the put. (Rubinstein)

'2) Simple chooser: Same tenor and strike for the call and the put.
'(Rubinstein)

'Features
'Chooser options are more expensive than standard options,
'since the purchaser has increased flexibility.

'Chooser options are path-independent.

'Benefits
'Chooser options provide the benefit of allowing hedging against both
'price increases and decreases without purchasing both a call and a put.

'Intrinsic Value Formula

'The payoff at expiration is:
'Max [C(Ec,Tc-t),P(Ec,Tc-t);t]
'Where C is the value of the call, P is the value of the put, Ec is the
'exercise or strike price of the call, Ep is the exercise or strike of
'the put, Tc is the tenor of the call, Tp is the tenor of the put, and t
'is the time until the choice is made. For a simple chooser,
'Ec = Ep and Tc = Tp.

'////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////
