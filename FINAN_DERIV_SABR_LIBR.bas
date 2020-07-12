Attribute VB_Name = "FINAN_DERIV_SABR_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : SABR_AUTO_CALL_OPTION_FUNC
'DESCRIPTION   : SABR MC BARRIER
'LIBRARY       : DERIVATIVES
'GROUP         : SABR
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function SABR_AUTO_CALL_OPTION_FUNC(ByVal SPOT As Double, _
ByVal COUPON As Double, _
ByVal KI_LEVEL As Double, _
ByVal REDEMPTION As Double, _
ByVal beta As Double, _
ByVal ALPHA As Double, _
ByVal RHO As Double, _
ByVal mu As Double, _
ByVal FWD As Double, _
ByVal RATE As Double, _
ByVal SETTLEMENT As Date, _
ByVal MATURITY As Date, _
Optional ByVal nSTEPS As Long = 100, _
Optional ByVal COUNT_BASIS As Double = 365, _
Optional ByVal VERSION As Integer = 0)

'FWD = SPOT / FORWARD

'Spot: 100
'Settlement: 02/09/2007
'Maturity:02/08/2010

'Coupon: 11%
'KI Level: 75%
'Redemption: 100%
'BETA: 0.9
'ALPHA: 0.3
'RHO: -0.1
'MU: -0.6
'FWD: 103.0328157 = SPOT / FORWARD 2010 = 100 / 0.970564565637823
'Sabr Atm Vol: 20.64%
'Steps: 10,000.00
'Black Scholes: 95.98%
'SABR: 96.66%
    
    Dim ii As Long
    Dim jj As Long
    
    Dim nDAYS As Long
    
    Dim FIRST_RND As Double
    Dim SECOND_RND As Double
    
    Dim NPV As Double
    Dim DELTA As Double
    
    Dim LOG_SPOT As Double
    Dim START_SIGMA As Double
    Dim ATM_SIGMA As Double

    Dim STEP_COUPON As Double
    Dim KI_FLAG As Boolean
    
    On Error GoTo ERROR_LABEL
    
    DELTA = 1 / COUNT_BASIS
    
    If VERSION <> 0 Then: beta = 1
    nDAYS = DateDiff("d", SETTLEMENT, MATURITY)
    ATM_SIGMA = SABR_ATM_OPTION_SIGMA_FUNC(ALPHA, beta, RHO, mu, FWD, nDAYS / COUNT_BASIS)
    
    NPV = 0
    
    For ii = 1 To nSTEPS
        LOG_SPOT = Log(SPOT)
        START_SIGMA = ATM_SIGMA
        KI_FLAG = False
        DELTA = 1 / COUNT_BASIS
        STEP_COUPON = COUPON
        
        For jj = 1 To nDAYS
        
            'Correlated random numbers
            FIRST_RND = NORMSINV_FUNC(Rnd, 0, 1, 0)
            SECOND_RND = RHO * FIRST_RND + _
                        Sqr(1 - RHO ^ 2) * _
                        NORMSINV_FUNC(Rnd, 0, 1, 0)
            
            'Simulate vol first
            START_SIGMA = START_SIGMA + mu * START_SIGMA * SECOND_RND * Sqr(DELTA)
            
            If VERSION <> 0 Then START_SIGMA = ATM_SIGMA
            
            'Simulate the SPOT
            LOG_SPOT = LOG_SPOT + (RATE - 0.5 * START_SIGMA ^ 2 * Exp(2 * _
                          (beta - 1) * (LOG_SPOT - RATE * _
                          (nDAYS - jj) / COUNT_BASIS))) * DELTA
            LOG_SPOT = LOG_SPOT + START_SIGMA * Exp((beta - 1) * _
                          (LOG_SPOT - RATE * (nDAYS - jj) / COUNT_BASIS)) _
                          * FIRST_RND * Sqr(DELTA)
            
            If Exp(LOG_SPOT) / SPOT < KI_LEVEL Then
                KI_FLAG = True
            End If
            
            If jj = COUNT_BASIS Then
                If Exp(LOG_SPOT) > SPOT Then
                    GoTo 1983
                End If
            ElseIf jj = (COUNT_BASIS * 2) Then
                If Exp(LOG_SPOT) > SPOT Then
                    STEP_COUPON = 2 * COUPON
                    GoTo 1983
                End If
            ElseIf jj = nDAYS Then
                If Exp(LOG_SPOT) > SPOT Then
                    STEP_COUPON = 3 * COUPON
                    GoTo 1983
                End If
            End If
            
            
        Next jj
        
        GoTo 1984
        
1983:
        NPV = NPV + 1 + STEP_COUPON
        GoTo 1985

1984:
        If KI_FLAG = False Then
            NPV = NPV + 1 * REDEMPTION
        Else
            NPV = NPV + 1 - IIf((1 - Exp(LOG_SPOT) / SPOT) < 0, _
                            1 - Exp(LOG_SPOT) / SPOT, 0)
        End If
1985:
'        Excel.Application.StatusBar = Format(100 * ii / nSTEPS, "0.00") & "% computed!"
    Next ii
    
    NPV = NPV / nSTEPS
    SABR_AUTO_CALL_OPTION_FUNC = NPV * Exp(-RATE * nDAYS / COUNT_BASIS)

Exit Function
ERROR_LABEL:
SABR_AUTO_CALL_OPTION_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : SABR_ATM_OPTION_SIGMA_FUNC
'DESCRIPTION   :
'LIBRARY       : DERIVATIVES
'GROUP         : SABR
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************


Public Function SABR_ATM_OPTION_SIGMA_FUNC(ByVal ALPHA As Double, _
ByVal beta As Double, _
ByVal RHO As Double, _
ByVal mu As Double, _
ByVal FWD As Double, _
ByVal EXPIRATION As Double)
      
    Dim TEMP_VAL As Double
      
    On Error GoTo ERROR_LABEL
                                
    TEMP_VAL = ((1 - beta) ^ 2) / 24
    TEMP_VAL = TEMP_VAL * (ALPHA ^ 2) / (FWD ^ (2 - 2 * beta))
    TEMP_VAL = TEMP_VAL + 0.25 * beta * RHO * mu * ALPHA / (FWD ^ (1 - beta))
    TEMP_VAL = TEMP_VAL + (2 - 3 * RHO ^ 2) * (mu ^ 2) / 24
    TEMP_VAL = TEMP_VAL * EXPIRATION
    TEMP_VAL = TEMP_VAL + 1
    TEMP_VAL = TEMP_VAL * ALPHA
    TEMP_VAL = TEMP_VAL / (FWD ^ (1 - beta))
    
    SABR_ATM_OPTION_SIGMA_FUNC = TEMP_VAL
    
Exit Function
ERROR_LABEL:
SABR_ATM_OPTION_SIGMA_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : SABR_OPTION_SIGMA_FUNC
'DESCRIPTION   :
'LIBRARY       : DERIVATIVES
'GROUP         : SABR
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function SABR_OPTION_SIGMA_FUNC(ByVal ALPHA As Double, _
ByVal beta As Double, _
ByVal RHO As Double, _
ByVal mu As Double, _
ByVal FWD As Double, _
ByVal EXPIRATION As Double, _
ByVal STRIKE As Double)
    
    Dim ZTEMP_VAL As Double
    Dim XTEMP_VAL As Double
    Dim YTEMP_VAL As Double
    Dim WTEMP_VAL As Double
    
    On Error GoTo ERROR_LABEL
    
    ZTEMP_VAL = mu / ALPHA
    ZTEMP_VAL = ZTEMP_VAL * (FWD * STRIKE) ^ ((1 - beta) / 2)
    ZTEMP_VAL = ZTEMP_VAL * Log(FWD / STRIKE)
    
    XTEMP_VAL = Log((Sqr(1 - 2 * RHO * ZTEMP_VAL + _
                ZTEMP_VAL ^ 2) + ZTEMP_VAL - RHO) / (1 - RHO))
    
    YTEMP_VAL = ALPHA ^ 2 * (1 / 24) * (1 - beta) ^ 2
    YTEMP_VAL = YTEMP_VAL / (FWD * STRIKE) ^ (1 - beta)
    YTEMP_VAL = YTEMP_VAL + 0.25 * RHO * beta * mu * ALPHA / _
            ((FWD * STRIKE) ^ ((1 - beta) / 2))
    YTEMP_VAL = YTEMP_VAL + (2 - 3 * RHO ^ 2) * mu ^ 2 / 24
    YTEMP_VAL = YTEMP_VAL * EXPIRATION + 1
    YTEMP_VAL = YTEMP_VAL * ALPHA
    YTEMP_VAL = YTEMP_VAL * ZTEMP_VAL
    YTEMP_VAL = YTEMP_VAL / XTEMP_VAL
    YTEMP_VAL = YTEMP_VAL / ((FWD * STRIKE) ^ ((1 - beta) / 2))
    
    WTEMP_VAL = ((1 - beta) ^ 2) / 24
    WTEMP_VAL = WTEMP_VAL * (Log(FWD / STRIKE)) ^ 2
    WTEMP_VAL = WTEMP_VAL + ((1 - beta) * Log(FWD / STRIKE)) ^ 4 / 1920 + 1
    
    SABR_OPTION_SIGMA_FUNC = YTEMP_VAL / WTEMP_VAL

Exit Function
ERROR_LABEL:
SABR_OPTION_SIGMA_FUNC = Err.number
End Function

'------------------------------------------------------------------------------------
'REFERENCE: Patrick S. Hagan et al, Managing Smile
'Risk, Wilmott magazine 021118_smile.pdf (siehe auch
'PatSmile.mnb) approximation of implied volatility,
'(2.17), p 89
'------------------------------------------------------------------------------------
