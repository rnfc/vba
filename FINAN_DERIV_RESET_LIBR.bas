Attribute VB_Name = "FINAN_DERIV_RESET_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : BINOMIAL_PLAIN_RESET_STRIKE_FUNC
'DESCRIPTION   : Haug and Haug Reset Binomial Model
'LIBRARY       : DERIVATIVES
'GROUP         : RESET
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function BINOMIAL_PLAIN_RESET_STRIKE_FUNC(ByVal SPOT As Double, _
ByVal STRIKE As Double, _
ByVal EXPIRATION As Double, _
ByVal RESET_TENOR As Double, _
ByVal RATE As Double, _
ByVal CARRY_COST As Double, _
ByVal VOLATILITY As Double, _
Optional ByVal POWER_FACTOR As Double = 1, _
Optional ByVal ALPHA As Double = 1, _
Optional ByVal nSTEPS As Long = 100, _
Optional ByVal OPTION_FLAG As Integer = -1)
        
'EXPIRATION: Time to Maturity
'RESET_TENOR: Years to Reset
'nSTEPS: Number of time steps
    
'POWER: S^Power-X^Power
'ALPHA: Reset in-out-of-money

    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim l As Long
    
    Dim ii As Long
    Dim jj As Long
                
    Dim U_VAL As Double
    Dim D_VAL As Double
    
    Dim DT_VAL As Double
    
    Dim AJ_VAL As Double
    Dim XR_VAL As Double
    
    Dim PROB_VAL As Double
    Dim START_VAL As Double
    Dim LAST_VAL As Double
    Dim OPTION_VAL As Double
    
    On Error GoTo ERROR_LABEL
    
    Select Case OPTION_FLAG
    Case 1 'Call option
        k = 1
    Case Else ' Put option
        k = -1
    End Select

    DT_VAL = EXPIRATION / nSTEPS
    U_VAL = Exp((CARRY_COST - VOLATILITY ^ 2 / 2) * DT_VAL + _
            VOLATILITY * Sqr(DT_VAL))
    D_VAL = Exp((CARRY_COST - VOLATILITY ^ 2 / 2) * DT_VAL - _
            VOLATILITY * Sqr(DT_VAL))
    
    l = Int(RESET_TENOR / DT_VAL)
    
    OPTION_VAL = 0
    For j = 0 To l
    
        START_VAL = SPOT * U_VAL ^ j * D_VAL ^ (l - j)
            
        If k = 1 Then ' Reset strike call
                XR_VAL = MINIMUM_FUNC(STRIKE, ALPHA * START_VAL)
        Else 'Reset strike put
                XR_VAL = MAXIMUM_FUNC(STRIKE, ALPHA * START_VAL)
        End If
        
         AJ_VAL = Int(Log(XR_VAL / (SPOT * U_VAL ^ j * D_VAL ^ _
                (nSTEPS - j))) / Log(U_VAL / D_VAL))
         
        If k = 1 Then   'call
            ii = MAXIMUM_FUNC(0, AJ_VAL + 1) + j
            jj = nSTEPS - l + j
        Else 'put
            ii = j
            jj = MINIMUM_FUNC(MAXIMUM_FUNC(AJ_VAL, 0) + j, nSTEPS - l + j)
        End If
        
        For i = ii To jj
            LAST_VAL = SPOT * U_VAL ^ i * D_VAL ^ (nSTEPS - i)
            PROB_VAL = FACTORIAL_FUNC(l) * FACTORIAL_FUNC(nSTEPS - l) / _
                      (FACTORIAL_FUNC(j) * FACTORIAL_FUNC(l - j) * _
                       FACTORIAL_FUNC(i - j) * _
                       FACTORIAL_FUNC(nSTEPS - l - i + j)) * 0.5 ^ nSTEPS
            
            OPTION_VAL = OPTION_VAL + PROB_VAL * k * (LAST_VAL ^ _
                          POWER_FACTOR - XR_VAL ^ POWER_FACTOR)
        Next i
    Next
    
    BINOMIAL_PLAIN_RESET_STRIKE_FUNC = Exp(-RATE * EXPIRATION) * OPTION_VAL

Exit Function
ERROR_LABEL:
BINOMIAL_PLAIN_RESET_STRIKE_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : BINOMIAL_PLAIN_RESET_STRIKE_BARRIER_FUNC
'DESCRIPTION   : Haug and Haug Reset Strike Binomial Model with Barrier
'LIBRARY       : DERIVATIVES
'GROUP         : RESET
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function BINOMIAL_PLAIN_RESET_STRIKE_BARRIER_FUNC( _
ByVal SPOT As Double, _
ByVal STRIKE As Double, _
ByVal BARRIER As Double, _
ByVal EXPIRATION As Double, _
ByVal RESET_TENOR As Double, _
ByVal RATE As Double, _
ByVal CARRY_COST As Double, _
ByVal VOLATILITY As Double, _
Optional ByVal POWER_FACTOR As Double = 1, _
Optional ByVal ALPHA As Double = 1, _
Optional ByVal nSTEPS As Long = 100, _
Optional ByVal OPTION_FLAG As Integer = -1, _
Optional ByVal COMPOUNDING_TYPE As Integer = 2)

    'EXPIRATION: Time to Maturity
    'RESET_TENOR: Years to Reset
    'nSTEPS: Number of time steps
    
    'POWER: S^Power-X^Power
    'ALPHA: Reset in-out-of-money
                
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim l As Long
    
    Dim ii As Long
    Dim jj As Long

    Dim U_VAL As Double
    Dim D_VAL As Double
    
    Dim HM_VAL As Double
    Dim HN_VAL As Double
    
    Dim XR_VAL As Double
    Dim PROB_VAL As Double
    
    Dim DT_VAL As Double
    
    Dim START_VAL As Double
    Dim LAST_VAL As Double
    
    Dim AJ_VAL As Double
    Dim OPTION_VAL As Double
    
    Dim HIT_PROB_VAL As Double 'Barrier Hit Prob
    Dim COMPOUNDING_FACTOR As Double

    On Error GoTo ERROR_LABEL

    Select Case OPTION_FLAG
    Case 1 'Call option
        k = 1
    Case Else ' Put option
        k = -1
    End Select
    
    Select Case COMPOUNDING_TYPE
    Case 1 'Continuous
        COMPOUNDING_FACTOR = 0
    Case 2 'Daily
        COMPOUNDING_FACTOR = 1 / 365
    Case 3 'Weekly
        COMPOUNDING_FACTOR = 1 / 52
    Case Else
        COMPOUNDING_FACTOR = 0 'Continuous
    End Select
    
    'Discrete barrier monitoring adjustment
    If BARRIER > SPOT Then
        BARRIER = BARRIER * Exp(0.5826 * VOLATILITY * Sqr(COMPOUNDING_FACTOR))
    ElseIf BARRIER < SPOT Then
        BARRIER = BARRIER * Exp(-0.5826 * VOLATILITY * Sqr(COMPOUNDING_FACTOR))
    End If

    DT_VAL = EXPIRATION / nSTEPS
    U_VAL = Exp((CARRY_COST - VOLATILITY ^ 2 / 2) * DT_VAL + _
            VOLATILITY * Sqr(DT_VAL))
    D_VAL = Exp((CARRY_COST - VOLATILITY ^ 2 / 2) * DT_VAL - _
            VOLATILITY * Sqr(DT_VAL))
    l = Int(RESET_TENOR / DT_VAL)
    
    OPTION_VAL = 0
    For j = 0 To l
    
        START_VAL = SPOT * U_VAL ^ j * D_VAL ^ (l - j)
            
        If k = 1 Then ' Reset strike call
                XR_VAL = MINIMUM_FUNC(STRIKE, ALPHA * START_VAL)
        Else 'Reset strike put
                XR_VAL = MAXIMUM_FUNC(STRIKE, ALPHA * START_VAL)
        End If
        
         AJ_VAL = Int(Log(XR_VAL / (SPOT * U_VAL ^ j * D_VAL ^ _
                (nSTEPS - j))) / Log(U_VAL / D_VAL))
         
        If k = 1 Then   'call
            ii = MAXIMUM_FUNC(0, AJ_VAL + 1) + j
            jj = nSTEPS - l + j
        Else 'put
            ii = j
            jj = MINIMUM_FUNC(MAXIMUM_FUNC(AJ_VAL, 0) + j, nSTEPS - l + j)
        End If
        
        For i = ii To jj
            
            LAST_VAL = SPOT * U_VAL ^ i * D_VAL ^ (nSTEPS - i)
            HM_VAL = Exp(-2 / (VOLATILITY ^ 2 * (l * DT_VAL)) * _
                    Abs(Log(SPOT / BARRIER) * Log(START_VAL / BARRIER)))
            HN_VAL = Exp(-2 / (VOLATILITY ^ 2 * (EXPIRATION - l * DT_VAL)) * _
                Abs(Log(START_VAL / BARRIER) * Log(LAST_VAL / BARRIER)))
            If SPOT > BARRIER Then ' Down and out
                If LAST_VAL <= BARRIER Then
                ' Probability of hitting the lower barrier
                    HN_VAL = 1
                ElseIf START_VAL <= BARRIER Then
                    HM_VAL = 1
                End If
            ElseIf SPOT < BARRIER Then ' Up and out
                If LAST_VAL >= BARRIER Then
                ' Probability of hitting the lower barrier
                    HN_VAL = 1
                ElseIf START_VAL >= BARRIER Then
                    HM_VAL = 1
                End If
            End If
            
            HIT_PROB_VAL = HM_VAL + HN_VAL - HM_VAL * HN_VAL

            PROB_VAL = (FACTORIAL_FUNC(l) * _
                        FACTORIAL_FUNC(nSTEPS - l)) / _
                       (FACTORIAL_FUNC(j) * _
                        FACTORIAL_FUNC(l - j) * _
                        FACTORIAL_FUNC(i - j) * _
                        FACTORIAL_FUNC(nSTEPS - l - i + j)) * 0.5 ^ nSTEPS
            
            OPTION_VAL = OPTION_VAL + PROB_VAL * k * _
                        (LAST_VAL ^ POWER_FACTOR - XR_VAL ^ POWER_FACTOR) * _
                        (1 - HIT_PROB_VAL)
    
        Next i
    Next j
    
    BINOMIAL_PLAIN_RESET_STRIKE_BARRIER_FUNC = _
        Exp(-RATE * EXPIRATION) * OPTION_VAL

Exit Function
ERROR_LABEL:
BINOMIAL_PLAIN_RESET_STRIKE_BARRIER_FUNC = Err.number
End Function
