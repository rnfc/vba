Attribute VB_Name = "FINAN_DERIV_BARRIER_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : STANDARD_BARRIER_OPTION_FUNC
'DESCRIPTION   : Standard barrier options
'LIBRARY       : DERIVATIVES
'GROUP         : BARRIER
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function STANDARD_BARRIER_OPTION_FUNC(ByVal SPOT As Double, _
ByVal STRIKE As Double, _
ByVal BARRIER As Double, _
ByVal cash As Double, _
ByVal TENOR As Double, _
ByVal RATE As Double, _
ByVal CARRY_COST As Double, _
ByVal SIGMA As Double, _
Optional ByVal OPTION_FLAG As Integer = 1, _
Optional ByVal MONITOR_BASIS As Integer = 2, _
Optional ByVal ADJ_FACTOR As Double = 0.5826, _
Optional ByVal CND_TYPE As Integer = 0)

'CASH: CASH REBATE

    Dim MU_VAL As Double
    Dim LAMBDA As Double
    
    Dim X1_VAL As Double
    Dim X2_VAL As Double
    
    Dim Y1_VAL As Double
    Dim Y2_VAL As Double
    
    Dim Z_VAL As Double
    
    Dim A_VAL As Double
    Dim B_VAL As Double
    Dim C_VAL As Double
    Dim D_VAL As Double
    Dim E_VAL As Double
    Dim F_VAL As Double

    Dim ETA_VAL As Integer
    Dim PHI_VAL As Integer
    
    On Error GoTo ERROR_LABEL

    BARRIER = DISCRETE_BARRIER_MONITORING_ADJ_FUNC(SPOT, BARRIER, SIGMA, _
    BARRIER_MONITORING_COUNT_BASIS_FUNC(MONITOR_BASIS), ADJ_FACTOR)

    MU_VAL = (CARRY_COST - SIGMA ^ 2 / 2) / SIGMA ^ 2
    LAMBDA = Sqr(MU_VAL ^ 2 + 2 * RATE / SIGMA ^ 2)
    
    X1_VAL = Log(SPOT / STRIKE) / (SIGMA * Sqr(TENOR)) + (1 + MU_VAL) * SIGMA * Sqr(TENOR)
    X2_VAL = Log(SPOT / BARRIER) / (SIGMA * Sqr(TENOR)) + (1 + MU_VAL) * SIGMA * Sqr(TENOR)
    
    Y1_VAL = Log(BARRIER ^ 2 / (SPOT * STRIKE)) / (SIGMA * Sqr(TENOR)) + _
    (1 + MU_VAL) * SIGMA * Sqr(TENOR)
    
    Y2_VAL = Log(BARRIER / SPOT) / (SIGMA * Sqr(TENOR)) + (1 + MU_VAL) * _
    SIGMA * Sqr(TENOR)
    
    Z_VAL = Log(BARRIER / SPOT) / (SIGMA * Sqr(TENOR)) + LAMBDA * _
    SIGMA * Sqr(TENOR)
        
            Select Case OPTION_FLAG
                Case 1, 5
                    ETA_VAL = 1
                    PHI_VAL = 1
                Case 2, 6
                    ETA_VAL = -1
                    PHI_VAL = 1
                Case 3, 7
                    ETA_VAL = 1
                    PHI_VAL = -1
                Case Else '4, 8
                    ETA_VAL = -1
                    PHI_VAL = -1
            End Select
    
'-----------------------------------------------------------
    '1: Down-and-in call    cdi
    '2: Up-and-in call  cui
    '3: Down-and-in put pdi
    '4: Up-and-in put   pui
    '5: Down-and-out call   cdo
    '6: Up-and-out call cuo
    '7: Down-and-out put    pdo
    '8: Up-and-out put  puo
'-----------------------------------------------------------
    
    A_VAL = PHI_VAL * SPOT * Exp((CARRY_COST - RATE) * TENOR) * CND_FUNC(PHI_VAL * X1_VAL, _
        CND_TYPE) - PHI_VAL * STRIKE * Exp(-RATE * TENOR) * CND_FUNC(PHI_VAL * X1_VAL - PHI_VAL _
        * SIGMA * Sqr(TENOR), CND_TYPE)
    
    B_VAL = PHI_VAL * SPOT * Exp((CARRY_COST - RATE) * TENOR) * CND_FUNC(PHI_VAL * X2_VAL, CND_TYPE) _
        - PHI_VAL * STRIKE * Exp(-RATE * TENOR) * CND_FUNC(PHI_VAL * X2_VAL - PHI_VAL * SIGMA * _
        Sqr(TENOR), CND_TYPE)
    
    C_VAL = PHI_VAL * SPOT * Exp((CARRY_COST - RATE) * TENOR) * (BARRIER / SPOT) ^ _
        (2 * (MU_VAL + 1)) * CND_FUNC(ETA_VAL * Y1_VAL, CND_TYPE) - PHI_VAL * STRIKE * Exp(-RATE * _
        TENOR) * (BARRIER / SPOT) ^ (2 * MU_VAL) * CND_FUNC(ETA_VAL * Y1_VAL - ETA_VAL * SIGMA * _
        Sqr(TENOR), CND_TYPE)
    
    D_VAL = PHI_VAL * SPOT * Exp((CARRY_COST - RATE) * TENOR) * (BARRIER / SPOT) ^ _
        (2 * (MU_VAL + 1)) * CND_FUNC(ETA_VAL * Y2_VAL, CND_TYPE) - PHI_VAL * STRIKE * Exp(-RATE * _
        TENOR) * (BARRIER / SPOT) ^ (2 * MU_VAL) * CND_FUNC(ETA_VAL * Y2_VAL - ETA_VAL * SIGMA * _
        Sqr(TENOR), CND_TYPE)
    
    E_VAL = cash * Exp(-RATE * TENOR) * (CND_FUNC(ETA_VAL * X2_VAL - ETA_VAL * SIGMA * Sqr(TENOR), _
        CND_TYPE) - (BARRIER / SPOT) ^ (2 * MU_VAL) * CND_FUNC(ETA_VAL * Y2_VAL - ETA_VAL * SIGMA * _
        Sqr(TENOR), CND_TYPE))
    
    F_VAL = cash * ((BARRIER / SPOT) ^ (MU_VAL + LAMBDA) * CND_FUNC(ETA_VAL * Z_VAL, CND_TYPE) + _
        (BARRIER / SPOT) ^ (MU_VAL - LAMBDA) * CND_FUNC(ETA_VAL * Z_VAL - 2 * ETA_VAL * LAMBDA * _
        SIGMA * Sqr(TENOR), CND_TYPE))
    

    If STRIKE > BARRIER Then
        
        Select Case OPTION_FLAG
            Case 1 ', "cdi" 'Down-and-in call
                STANDARD_BARRIER_OPTION_FUNC = C_VAL + E_VAL
            Case 2 ', "cui" 'Up-and-in call
                STANDARD_BARRIER_OPTION_FUNC = A_VAL + E_VAL
            Case 3 ', "pdi" 'Down-and-in put
                STANDARD_BARRIER_OPTION_FUNC = B_VAL - C_VAL + D_VAL + E_VAL
            Case 4 ', "pui" 'Up-and-in put
                STANDARD_BARRIER_OPTION_FUNC = A_VAL - B_VAL + D_VAL + E_VAL
            Case 5 ', "cdo" 'Down-and-out call
                STANDARD_BARRIER_OPTION_FUNC = A_VAL - C_VAL + F_VAL
            Case 6 ', "cuo" 'Up-and-out call
                STANDARD_BARRIER_OPTION_FUNC = F_VAL
            Case 7 ', "pdo" 'Down-and-out put
                STANDARD_BARRIER_OPTION_FUNC = A_VAL - B_VAL + C_VAL - D_VAL + F_VAL
            Case Else '8, "puo" 'Up-and-out put
                STANDARD_BARRIER_OPTION_FUNC = B_VAL - D_VAL + F_VAL
            End Select
    
    ElseIf STRIKE < BARRIER Then
        
        Select Case OPTION_FLAG
            Case 1 ', "cdi" 'Down-and-in call
                STANDARD_BARRIER_OPTION_FUNC = A_VAL - B_VAL + D_VAL + E_VAL
            Case 2 ', "cui" 'Up-and-in call
                STANDARD_BARRIER_OPTION_FUNC = B_VAL - C_VAL + D_VAL + E_VAL
            Case 3 ', "pdi" 'Down-and-in put
                STANDARD_BARRIER_OPTION_FUNC = A_VAL + E_VAL
            Case 4 ', "pui" 'Up-and-in put
                STANDARD_BARRIER_OPTION_FUNC = C_VAL + E_VAL
            Case 5 ', "cdo" 'Down-and-out call
                STANDARD_BARRIER_OPTION_FUNC = B_VAL + F_VAL - D_VAL
            Case 6 ', "cuo" 'Up-and-out call
                STANDARD_BARRIER_OPTION_FUNC = A_VAL - B_VAL + C_VAL - D_VAL + F_VAL
            Case 7 ', "pdo" 'Down-and-out put
                STANDARD_BARRIER_OPTION_FUNC = F_VAL
            Case Else '8, "puo" 'Up-and-out put
                STANDARD_BARRIER_OPTION_FUNC = A_VAL - C_VAL + F_VAL
        End Select
    
    End If
    
Exit Function
ERROR_LABEL:
STANDARD_BARRIER_OPTION_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : TWO_ASSET_BARRIER_OPTION_FUNC
'DESCRIPTION   : Two asset barrier options
'LIBRARY       : DERIVATIVES
'GROUP         : BARRIER
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009

'************************************************************************************
'************************************************************************************

Function TWO_ASSET_BARRIER_OPTION_FUNC(ByVal SPOT_A As Double, _
ByVal SPOT_B As Double, _
ByVal STRIKE As Double, _
ByVal BARRIER As Double, _
ByVal TENOR As Double, _
ByVal RATE As Double, _
ByVal CARRY_COST_A As Double, _
ByVal CARRY_COST_B As Double, _
ByVal SIGMA_A As Double, _
ByVal SIGMA_B As Double, _
ByVal RHO_VAL As Double, _
Optional ByVal OPTION_FLAG As Integer = 1, _
Optional ByVal MONITOR_BASIS As Integer = 2, _
Optional ByVal ADJ_FACTOR As Double = 0.5826, _
Optional ByVal CND_TYPE As Integer = 0, _
Optional ByVal CBND_TYPE As Integer = 0)
 
'OPTION_FLAG Options:
 
   '1: Down-and-in call --> cdi
   '2: Up-and-in call --> cui
   '3: Down-and-in put --> pdi
   '4: Up-and-in put --> pui
   '5: Down-and-out call --> cdo
   '6: Up-and-out call --> cuo
   '7: Down-and-out put --> pdo
   '8: Up-and-out put --> puo
 
    Dim D1_VAL As Double
    Dim D2_VAL As Double
    Dim D3_VAL As Double
    Dim D4_VAL As Double
    
    Dim E1_VAL As Double
    Dim E2_VAL As Double
    Dim E3_VAL As Double
    Dim E4_VAL As Double
    
    Dim BMU_VAL As Double
    Dim AMU_VAL As Double
    
    Dim ETA_VAL As Integer    '1 for call options and -1 for put options
    Dim PHI_VAL As Integer    '1 for up options and -1 for down options
    
    Dim KNOCK_OUT_VAL As Double
    
    On Error GoTo ERROR_LABEL
   
    BARRIER = DISCRETE_BARRIER_MONITORING_ADJ_FUNC(SPOT_B, BARRIER, SIGMA_B, _
    BARRIER_MONITORING_COUNT_BASIS_FUNC(MONITOR_BASIS), ADJ_FACTOR)

    BMU_VAL = CARRY_COST_A - SIGMA_A ^ 2 / 2
    AMU_VAL = CARRY_COST_B - SIGMA_B ^ 2 / 2
    
    D1_VAL = (Log(SPOT_A / STRIKE) + (BMU_VAL + SIGMA_A ^ 2 / 2) * TENOR) / _
    (SIGMA_A * Sqr(TENOR))
    
    D2_VAL = D1_VAL - SIGMA_A * Sqr(TENOR)
    D3_VAL = D1_VAL + 2 * RHO_VAL * Log(BARRIER / SPOT_B) / (SIGMA_B * Sqr(TENOR))
    D4_VAL = D2_VAL + 2 * RHO_VAL * Log(BARRIER / SPOT_B) / (SIGMA_B * Sqr(TENOR))
    
    E1_VAL = (Log(BARRIER / SPOT_B) - (AMU_VAL + RHO_VAL * SIGMA_A * SIGMA_B) * TENOR) / _
    (SIGMA_B * Sqr(TENOR))
    
    E2_VAL = E1_VAL + RHO_VAL * SIGMA_A * Sqr(TENOR)
    E3_VAL = E1_VAL - 2 * Log(BARRIER / SPOT_B) / (SIGMA_B * Sqr(TENOR))
    E4_VAL = E2_VAL - 2 * Log(BARRIER / SPOT_B) / (SIGMA_B * Sqr(TENOR))
  
        Select Case OPTION_FLAG
            Case 6, 2 ' "cuo", "cui"
                ETA_VAL = 1
                PHI_VAL = 1
            Case 5, 1 '"cdo", "cdi"
                ETA_VAL = 1
                PHI_VAL = -1
            Case 8, 4 '"puo", "pui"
                ETA_VAL = -1
                PHI_VAL = 1
            Case Else '7, 3 '"pdo", "pdi"
                ETA_VAL = -1
                PHI_VAL = -1
        End Select
            
    KNOCK_OUT_VAL = ETA_VAL * SPOT_A * Exp((CARRY_COST_A - RATE) * TENOR) * _
        (CBND_FUNC(ETA_VAL * D1_VAL, PHI_VAL * E1_VAL, -ETA_VAL * PHI_VAL * RHO_VAL, CND_TYPE, CBND_TYPE) - _
        Exp(2 * (AMU_VAL + RHO_VAL * SIGMA_A * SIGMA_B) * Log(BARRIER / SPOT_B) / _
        SIGMA_B ^ 2) * CBND_FUNC(ETA_VAL * D3_VAL, PHI_VAL * E3_VAL, -ETA_VAL * PHI_VAL * RHO_VAL, CND_TYPE, _
        CBND_TYPE)) - ETA_VAL * Exp(-RATE * TENOR) * STRIKE * (CBND_FUNC(ETA_VAL * D2_VAL, PHI_VAL * _
        E2_VAL, -ETA_VAL * PHI_VAL * RHO_VAL, CND_TYPE, CBND_TYPE) - Exp(2 * AMU_VAL * Log(BARRIER / _
        SPOT_B) / SIGMA_B ^ 2) * CBND_FUNC(ETA_VAL * D4_VAL, PHI_VAL * E4_VAL, -ETA_VAL * PHI_VAL * RHO_VAL, _
        CND_TYPE, CBND_TYPE))


    Select Case OPTION_FLAG
        Case 5, 6, 7, 8 ' "cdo", "cuo", "pdo", "puo"
            TWO_ASSET_BARRIER_OPTION_FUNC = KNOCK_OUT_VAL
        Case 2, 1 '"cui", "cdi"
            TWO_ASSET_BARRIER_OPTION_FUNC = GENERALIZED_BLACK_SCHOLES_FUNC(SPOT_A, STRIKE, TENOR, _
            RATE, CARRY_COST_A, SIGMA_A, 1, CND_TYPE) - KNOCK_OUT_VAL
        Case Else '4, 3 '"pui", "pdi"
            TWO_ASSET_BARRIER_OPTION_FUNC = GENERALIZED_BLACK_SCHOLES_FUNC(SPOT_A, STRIKE, TENOR, _
            RATE, CARRY_COST_A, SIGMA_A, -1, CND_TYPE) - KNOCK_OUT_VAL
    End Select
    
Exit Function
ERROR_LABEL:
TWO_ASSET_BARRIER_OPTION_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : DOUBLE_BARRIER_OPTION_FUNC
'DESCRIPTION   : Double barrier options
'LIBRARY       : DERIVATIVES
'GROUP         : BARRIER
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009

'************************************************************************************
'************************************************************************************

Function DOUBLE_BARRIER_OPTION_FUNC(ByVal SPOT As Double, _
ByVal STRIKE As Double, _
ByVal LOWER_BARRIER As Double, _
ByVal UPPER_BARRIER As Double, _
ByVal TENOR As Double, _
ByVal RATE As Double, _
ByVal CARRY_COST As Double, _
ByVal SIGMA As Double, _
ByVal UPPER_CURVE As Double, _
ByVal LOWER_CURVE As Double, _
Optional ByVal OPTION_FLAG As Integer = 1, _
Optional ByVal MONITOR_BASIS As Integer = 2, _
Optional ByVal ADJ_FACTOR As Double = 0.5826, _
Optional ByVal CND_TYPE As Integer = 0)

'Upper curvature (delta_a)
'Lower curvature(delta_b)
    
    Dim i As Integer
    Dim SROW As Integer
    Dim NROWS As Integer
    
    Dim E_VAL As Double
    Dim F_VAL As Double
    
    Dim D1_VAL As Double
    Dim D2_VAL As Double
    
    Dim D3_VAL As Double
    Dim D4_VAL As Double
    
    Dim MU1_VAL As Double
    Dim MU2_VAL As Double
    Dim MU3_VAL As Double
    
    Dim TEMP_VAL As Double

    Dim ATEMP_SUM As Double
    Dim BTEMP_SUM As Double

    On Error GoTo ERROR_LABEL

    UPPER_BARRIER = DISCRETE_BARRIER_MONITORING_ADJ_FUNC(SPOT, UPPER_BARRIER, SIGMA, _
    BARRIER_MONITORING_COUNT_BASIS_FUNC(MONITOR_BASIS), ADJ_FACTOR)

    LOWER_BARRIER = DISCRETE_BARRIER_MONITORING_ADJ_FUNC(SPOT, LOWER_BARRIER, SIGMA, _
    BARRIER_MONITORING_COUNT_BASIS_FUNC(MONITOR_BASIS), ADJ_FACTOR)

    F_VAL = UPPER_BARRIER * Exp(UPPER_CURVE * TENOR)
    E_VAL = LOWER_BARRIER * Exp(UPPER_CURVE * TENOR)
    
    ATEMP_SUM = 0
    BTEMP_SUM = 0
    
    SROW = -5
    NROWS = 5

'1: Call up-and-out-down-and-out
'2: Put up-and-out-down-and-out
'3: Call up-and-in-down-and-in
'4: Put up-and-in-down-and-in
    
 Select Case OPTION_FLAG
     Case 1, 3 ', "co", "ci"
        
        For i = SROW To NROWS
            
            D1_VAL = (Log(SPOT * UPPER_BARRIER ^ (2 * i) / (STRIKE * LOWER_BARRIER ^ _
                (2 * i))) + (CARRY_COST + SIGMA ^ 2 / 2) * TENOR) / (SIGMA * _
                Sqr(TENOR))
            
            D2_VAL = (Log(SPOT * UPPER_BARRIER ^ (2 * i) / (F_VAL * LOWER_BARRIER ^ _
                (2 * i))) + (CARRY_COST + SIGMA ^ 2 / 2) * TENOR) / (SIGMA * _
                Sqr(TENOR))
            
            D3_VAL = (Log(LOWER_BARRIER ^ (2 * i + 2) / (STRIKE * SPOT * _
                UPPER_BARRIER ^ (2 * i))) + (CARRY_COST + SIGMA ^ 2 / 2) * _
                TENOR) / (SIGMA * Sqr(TENOR))
            
            D4_VAL = (Log(LOWER_BARRIER ^ (2 * i + 2) / (F_VAL * SPOT * UPPER_BARRIER _
                ^ (2 * i))) + (CARRY_COST + SIGMA ^ 2 / 2) * TENOR) / (SIGMA * _
                Sqr(TENOR))
            
            MU1_VAL = 2 * (CARRY_COST - LOWER_CURVE - i * (UPPER_CURVE - LOWER_CURVE)) / _
            SIGMA ^ 2 + 1
            
            MU2_VAL = 2 * i * (UPPER_CURVE - LOWER_CURVE) / SIGMA ^ 2
            
            MU3_VAL = 2 * (CARRY_COST - LOWER_CURVE + i * (UPPER_CURVE - LOWER_CURVE)) / _
            SIGMA ^ 2 + 1
            
            ATEMP_SUM = ATEMP_SUM + (UPPER_BARRIER ^ i / LOWER_BARRIER ^ i) ^ MU1_VAL * _
                (LOWER_BARRIER / SPOT) ^ MU2_VAL * (CND_FUNC(D1_VAL, CND_TYPE) - CND_FUNC(D2_VAL, _
                CND_TYPE)) - (LOWER_BARRIER ^ (i + 1) / (UPPER_BARRIER ^ i * _
                SPOT)) ^ MU3_VAL * (CND_FUNC(D3_VAL, CND_TYPE) - CND_FUNC(D4_VAL, CND_TYPE))
            
            BTEMP_SUM = BTEMP_SUM + (UPPER_BARRIER ^ i / LOWER_BARRIER ^ i) ^ _
                (MU1_VAL - 2) * (LOWER_BARRIER / SPOT) ^ MU2_VAL * (CND_FUNC(D1_VAL - SIGMA * _
                Sqr(TENOR), CND_TYPE) - CND_FUNC(D2_VAL - SIGMA * Sqr(TENOR), CND_TYPE)) _
                - (LOWER_BARRIER ^ (i + 1) / (UPPER_BARRIER ^ i * SPOT)) ^ _
                (MU3_VAL - 2) * (CND_FUNC(D3_VAL - SIGMA * Sqr(TENOR), CND_TYPE) - _
                CND_FUNC(D4_VAL - SIGMA * Sqr(TENOR), CND_TYPE))
        Next i
        
        TEMP_VAL = SPOT * Exp((CARRY_COST - RATE) * TENOR) * ATEMP_SUM - STRIKE * _
            Exp(-RATE * TENOR) * BTEMP_SUM
     
     Case Else '2, 4 ', "po", "pi"
        
        For i = SROW To NROWS
            
            D1_VAL = (Log(SPOT * UPPER_BARRIER ^ (2 * i) / (E_VAL * LOWER_BARRIER ^ _
                (2 * i))) + (CARRY_COST + SIGMA ^ 2 / 2) * TENOR) / (SIGMA * _
                Sqr(TENOR))
            
            D2_VAL = (Log(SPOT * UPPER_BARRIER ^ (2 * i) / (STRIKE * _
                LOWER_BARRIER ^ (2 * i))) + (CARRY_COST + SIGMA ^ 2 / 2) * _
                TENOR) / (SIGMA * Sqr(TENOR))
            
            D3_VAL = (Log(LOWER_BARRIER ^ (2 * i + 2) / (E_VAL * SPOT * UPPER_BARRIER ^ _
                (2 * i))) + (CARRY_COST + SIGMA ^ 2 / 2) * TENOR) / (SIGMA * _
                Sqr(TENOR))
            
            D4_VAL = (Log(LOWER_BARRIER ^ (2 * i + 2) / (STRIKE * SPOT * UPPER_BARRIER ^ _
            (2 * i))) + (CARRY_COST + SIGMA ^ 2 / 2) * TENOR) / (SIGMA * Sqr(TENOR))
            
            MU1_VAL = 2 * (CARRY_COST - LOWER_CURVE - i * (UPPER_CURVE - LOWER_CURVE)) / _
            SIGMA ^ 2 + 1
            
            MU2_VAL = 2 * i * (UPPER_CURVE - LOWER_CURVE) / SIGMA ^ 2
            
            MU3_VAL = 2 * (CARRY_COST - LOWER_CURVE + i * (UPPER_CURVE - LOWER_CURVE)) / _
            SIGMA ^ 2 + 1
            
            ATEMP_SUM = ATEMP_SUM + (UPPER_BARRIER ^ i / LOWER_BARRIER ^ i) ^ MU1_VAL * _
                (LOWER_BARRIER / SPOT) ^ MU2_VAL * (CND_FUNC(D1_VAL, CND_TYPE) - CND_FUNC(D2_VAL, _
                CND_TYPE)) - (LOWER_BARRIER ^ (i + 1) / (UPPER_BARRIER ^ i * _
                SPOT)) ^ MU3_VAL * (CND_FUNC(D3_VAL, CND_TYPE) - CND_FUNC(D4_VAL, CND_TYPE))
            
            BTEMP_SUM = BTEMP_SUM + (UPPER_BARRIER ^ i / LOWER_BARRIER ^ i) ^ _
                (MU1_VAL - 2) * (LOWER_BARRIER / SPOT) ^ MU2_VAL * (CND_FUNC(D1_VAL - SIGMA * _
                Sqr(TENOR), CND_TYPE) - CND_FUNC(D2_VAL - SIGMA * Sqr(TENOR), CND_TYPE)) _
                - (LOWER_BARRIER ^ (i + 1) / (UPPER_BARRIER ^ i * SPOT)) ^ _
                (MU3_VAL - 2) * (CND_FUNC(D3_VAL - SIGMA * Sqr(TENOR), CND_TYPE) - _
                CND_FUNC(D4_VAL - SIGMA * Sqr(TENOR), CND_TYPE))
        Next i
        
        TEMP_VAL = STRIKE * Exp(-RATE * TENOR) * BTEMP_SUM - SPOT * _
            Exp((CARRY_COST - RATE) * TENOR) * ATEMP_SUM
        
    End Select
    
    Select Case OPTION_FLAG
        Case 1, 2 ', "co", "po"
        '1: Call up-and-out-down-and-out
        '2: Put up-and-out-down-and-out
            DOUBLE_BARRIER_OPTION_FUNC = TEMP_VAL
        Case 3 ', "ci"
        '3: Call up-and-in-down-and-in
            DOUBLE_BARRIER_OPTION_FUNC = GENERALIZED_BLACK_SCHOLES_FUNC(SPOT, STRIKE, TENOR, RATE, _
            CARRY_COST, SIGMA, 1, CND_TYPE) - TEMP_VAL
        Case Else '4 ', "pi"
        '4: Put up-and-in-down-and-in
            DOUBLE_BARRIER_OPTION_FUNC = GENERALIZED_BLACK_SCHOLES_FUNC(SPOT, STRIKE, TENOR, RATE, _
            CARRY_COST, SIGMA, -1, CND_TYPE) - TEMP_VAL
    End Select
    
Exit Function
ERROR_LABEL:
DOUBLE_BARRIER_OPTION_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : PARTIAL_TIME_SINGLE_ASSET_BARRIER_OPTION_FUNC
'DESCRIPTION   : Partial-time single asset barrier options
'LIBRARY       : DERIVATIVES
'GROUP         : BARRIER
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009

'************************************************************************************
'************************************************************************************

Function PARTIAL_TIME_SINGLE_ASSET_BARRIER_OPTION_FUNC(ByVal SPOT As Double, _
ByVal STRIKE As Double, _
ByVal BARRIER As Double, _
ByVal FIRST_TENOR As Double, _
ByVal SECOND_TENOR As Double, _
ByVal RATE As Double, _
ByVal CARRY_COST As Double, _
ByVal SIGMA As Double, _
Optional ByVal OPTION_FLAG As Integer = 1, _
Optional ByVal MONITOR_BASIS As Integer = 2, _
Optional ByVal ADJ_FACTOR As Double = 0.5826, _
Optional ByVal CND_TYPE As Integer = 0, _
Optional ByVal CBND_TYPE As Integer = 0)
    
    Dim D1_VAL As Double
    Dim D2_VAL As Double
    
    Dim F1_VAL As Double
    Dim F2_VAL As Double
    
    Dim E1_VAL As Double
    Dim E2_VAL As Double
    
    Dim E3_VAL As Double
    Dim E4_VAL As Double
    
    Dim G1_VAL As Double
    Dim G2_VAL As Double
    
    Dim G3_VAL As Double
    Dim G4_VAL As Double
    
    Dim Z1_VAL As Double
    Dim Z2_VAL As Double
    Dim Z3_VAL As Double
    Dim Z4_VAL As Double
    Dim Z5_VAL As Double
    Dim Z6_VAL As Double
    Dim Z7_VAL As Double
    Dim Z8_VAL As Double
    
    Dim MU_VAL As Double
    Dim RHO_VAL As Double
    Dim ETA_VAL As Integer

    Dim TEMP_BARRIER As Double
    
    On Error GoTo ERROR_LABEL
    
    TEMP_BARRIER = DISCRETE_BARRIER_MONITORING_ADJ_FUNC(SPOT, BARRIER, SIGMA, _
    BARRIER_MONITORING_COUNT_BASIS_FUNC(MONITOR_BASIS), ADJ_FACTOR)
    
'1 Up-and-out call type a: (BARRIER>STRIKE)   --> cuoa
'2 Down-and-out call type a: (BARRIER<STRIKE) --> cdoa

'3 Up-and-out put type a: (BARRIER>STRIKE)    --> puoa
'4 Down-and-out put type a: (BARRIER<STRIKE)   --> pdoa

'5 Out call type b1: --> cob1
'6 Out put type b1: --> pob1

'7 Up-and-out call type b2: (BARRIER>STRIKE) --> cuob2
'8 Down-and-out call type b2: (BARRIER<STRIKE) --> cdob2

'9  Up-and-out put type b2: (BARRIER>STRIKE)   --> puob2
'10  Down-and-out put type b2: (BARRIER<STRIKE)  --> pdob2

        Select Case OPTION_FLAG
            Case 1 ', "cuoa"
                ETA_VAL = -1
            Case Else '2 ', "cdoa"
                ETA_VAL = 1
        End Select
    
    D1_VAL = (Log(SPOT / STRIKE) + (CARRY_COST + SIGMA ^ 2 / 2) * SECOND_TENOR) / _
    (SIGMA * Sqr(SECOND_TENOR))
    D2_VAL = D1_VAL - SIGMA * Sqr(SECOND_TENOR)
    
    F1_VAL = (Log(SPOT / STRIKE) + 2 * Log(TEMP_BARRIER / SPOT) + (CARRY_COST + _
        SIGMA ^ 2 / 2) * SECOND_TENOR) / (SIGMA * Sqr(SECOND_TENOR))
    F2_VAL = F1_VAL - SIGMA * Sqr(SECOND_TENOR)
    
    E1_VAL = (Log(SPOT / TEMP_BARRIER) + (CARRY_COST + SIGMA ^ 2 / 2) * FIRST_TENOR) / _
    (SIGMA * Sqr(FIRST_TENOR))
    E2_VAL = E1_VAL - SIGMA * Sqr(FIRST_TENOR)
    E3_VAL = E1_VAL + 2 * Log(TEMP_BARRIER / SPOT) / (SIGMA * Sqr(FIRST_TENOR))
    E4_VAL = E3_VAL - SIGMA * Sqr(FIRST_TENOR)
    
    MU_VAL = (CARRY_COST - SIGMA ^ 2 / 2) / SIGMA ^ 2
    RHO_VAL = Sqr(FIRST_TENOR / SECOND_TENOR)
    
    G1_VAL = (Log(SPOT / TEMP_BARRIER) + (CARRY_COST + SIGMA ^ 2 / 2) * SECOND_TENOR) / _
    (SIGMA * Sqr(SECOND_TENOR))
    G2_VAL = G1_VAL - SIGMA * Sqr(SECOND_TENOR)
    G3_VAL = G1_VAL + 2 * Log(TEMP_BARRIER / SPOT) / (SIGMA * Sqr(SECOND_TENOR))
    G4_VAL = G3_VAL - SIGMA * Sqr(SECOND_TENOR)
    
    Z1_VAL = CND_FUNC(E2_VAL, CND_TYPE) - (TEMP_BARRIER / SPOT) ^ (2 * MU_VAL) * CND_FUNC(E4_VAL, CND_TYPE)
    Z2_VAL = CND_FUNC(-E2_VAL, CND_TYPE) - (TEMP_BARRIER / SPOT) ^ (2 * MU_VAL) * CND_FUNC(-E4_VAL, CND_TYPE)
    Z3_VAL = CBND_FUNC(G2_VAL, E2_VAL, RHO_VAL, CND_TYPE, CBND_TYPE) - (TEMP_BARRIER / SPOT) ^ _
        (2 * MU_VAL) * CBND_FUNC(G4_VAL, -E4_VAL, -RHO_VAL, CND_TYPE, CBND_TYPE)
    Z4_VAL = CBND_FUNC(-G2_VAL, -E2_VAL, RHO_VAL, CND_TYPE, CBND_TYPE) - (TEMP_BARRIER / SPOT) ^ _
        (2 * MU_VAL) * CBND_FUNC(-G4_VAL, E4_VAL, -RHO_VAL, CND_TYPE, CBND_TYPE)
    Z5_VAL = CND_FUNC(E1_VAL, CND_TYPE) - (TEMP_BARRIER / SPOT) ^ (2 * (MU_VAL + 1)) * CND_FUNC(E3_VAL, _
        CND_TYPE)
    Z6_VAL = CND_FUNC(-E1_VAL, CND_TYPE) - (TEMP_BARRIER / SPOT) ^ (2 * (MU_VAL + 1)) * _
        CND_FUNC(-E3_VAL, CND_TYPE)
    Z7_VAL = CBND_FUNC(G1_VAL, E1_VAL, RHO_VAL, CND_TYPE, CBND_TYPE) - (TEMP_BARRIER / SPOT) ^ _
        (2 * (MU_VAL + 1)) * CBND_FUNC(G3_VAL, -E3_VAL, -RHO_VAL, CND_TYPE, CBND_TYPE)
    Z8_VAL = CBND_FUNC(-G1_VAL, -E1_VAL, RHO_VAL, CND_TYPE, CBND_TYPE) - (TEMP_BARRIER / SPOT) ^ _
        (2 * (MU_VAL + 1)) * CBND_FUNC(-G3_VAL, E3_VAL, -RHO_VAL, CND_TYPE, CBND_TYPE)
    


    '----------------------------------------------------------------------------
    '-----------------------call down-and out and up-and-out type a
    '----------------------------------------------------------------------------
    
    If ((OPTION_FLAG = 2) Or (OPTION_FLAG = 1)) Then
    'Or (OPTION_FLAG = "cdoa") Or (OPTION_FLAG = "cuoa")) Then 'PERFECT
        
        PARTIAL_TIME_SINGLE_ASSET_BARRIER_OPTION_FUNC = SPOT * Exp((CARRY_COST - RATE) * SECOND_TENOR) * _
            (CBND_FUNC(D1_VAL, ETA_VAL * E1_VAL, ETA_VAL * RHO_VAL, CND_TYPE, CBND_TYPE) - _
            (TEMP_BARRIER / SPOT) ^ (2 * (MU_VAL + 1)) * _
            CBND_FUNC(F1_VAL, ETA_VAL * E3_VAL, ETA_VAL * RHO_VAL, CND_TYPE, CBND_TYPE)) - _
            STRIKE * Exp(-RATE * SECOND_TENOR) * _
            (CBND_FUNC(D2_VAL, ETA_VAL * E2_VAL, ETA_VAL * RHO_VAL, CND_TYPE, CBND_TYPE) - _
            (TEMP_BARRIER / SPOT) ^ (2 * MU_VAL) * _
            CBND_FUNC(F2_VAL, ETA_VAL * E4_VAL, ETA_VAL * RHO_VAL, CND_TYPE, CBND_TYPE))
    
    '---------------------------------------------------------------------------
    '------------------------call down-and-out type b2
    '---------------------------------------------------------------------------
    
    ElseIf ((OPTION_FLAG = 8)) And (STRIKE < TEMP_BARRIER) Then
        '(OPTION_FLAG = "cdob2") Or
        PARTIAL_TIME_SINGLE_ASSET_BARRIER_OPTION_FUNC = SPOT * Exp((CARRY_COST - RATE) * SECOND_TENOR) * _
            (CBND_FUNC(G1_VAL, E1_VAL, RHO_VAL, CND_TYPE, CBND_TYPE) - (TEMP_BARRIER / SPOT) ^ _
            (2 * (MU_VAL + 1)) * CBND_FUNC(G3_VAL, -E3_VAL, -RHO_VAL, CND_TYPE, CBND_TYPE)) _
            - STRIKE * Exp(-RATE * SECOND_TENOR) * (CBND_FUNC(G2_VAL, E2_VAL, RHO_VAL, _
            CND_TYPE, CBND_TYPE) - (TEMP_BARRIER / SPOT) ^ _
            (2 * MU_VAL) * CBND_FUNC(G4_VAL, -E4_VAL, -RHO_VAL, CND_TYPE, CBND_TYPE))
    
    ElseIf ((OPTION_FLAG = 8)) And (STRIKE > TEMP_BARRIER) Then
    '(OPTION_FLAG = "cdob2") Or
        PARTIAL_TIME_SINGLE_ASSET_BARRIER_OPTION_FUNC = _
        PARTIAL_TIME_SINGLE_ASSET_BARRIER_OPTION_FUNC(SPOT, STRIKE, BARRIER, _
        FIRST_TENOR, SECOND_TENOR, RATE, CARRY_COST, SIGMA, 5, MONITOR_BASIS, _
        ADJ_FACTOR) 'Out call type b1
        
    
    '---------------------------------------------------------------------------
    '------------------------call up-and-out type b2
    '---------------------------------------------------------------------------
    
    ElseIf ((OPTION_FLAG = 7)) And (STRIKE < TEMP_BARRIER) Then
        '(OPTION_FLAG = "cuob2") Or
        PARTIAL_TIME_SINGLE_ASSET_BARRIER_OPTION_FUNC = SPOT * Exp((CARRY_COST - RATE) * SECOND_TENOR) * _
            (CBND_FUNC(-G1_VAL, -E1_VAL, RHO_VAL, CND_TYPE, CBND_TYPE) - (TEMP_BARRIER / _
            SPOT) ^ (2 * (MU_VAL + 1)) * CBND_FUNC(-G3_VAL, E3_VAL, -RHO_VAL, CND_TYPE, CBND_TYPE)) _
            - STRIKE * Exp(-RATE * SECOND_TENOR) * (CBND_FUNC(-G2_VAL, -E2_VAL, RHO_VAL, _
            CND_TYPE, CBND_TYPE) - (TEMP_BARRIER / SPOT) ^ _
            (2 * MU_VAL) * CBND_FUNC(-G4_VAL, E4_VAL, -RHO_VAL, CND_TYPE, CBND_TYPE)) - SPOT * _
            Exp((CARRY_COST - RATE) * SECOND_TENOR) _
            * (CBND_FUNC(-D1_VAL, -E1_VAL, RHO_VAL, CND_TYPE, CBND_TYPE) - (TEMP_BARRIER / _
            SPOT) ^ (2 * (MU_VAL + 1)) * CBND_FUNC(E3_VAL, -F1_VAL, -RHO_VAL, CND_TYPE, CBND_TYPE)) _
            + STRIKE * Exp(-RATE * SECOND_TENOR) * (CBND_FUNC(-D2_VAL, -E2_VAL, RHO_VAL, _
            CND_TYPE, CBND_TYPE) - (TEMP_BARRIER / SPOT) ^ _
            (2 * MU_VAL) * CBND_FUNC(E4_VAL, -F2_VAL, -RHO_VAL, CND_TYPE, CBND_TYPE))
    
    '---------------------------------------------------------------------------
    '------------------------call out type a1
    '---------------------------------------------------------------------------
    
    ElseIf ((OPTION_FLAG = 5)) And (STRIKE > TEMP_BARRIER) Then
        '(OPTION_FLAG = "cob1") Or
        PARTIAL_TIME_SINGLE_ASSET_BARRIER_OPTION_FUNC = SPOT * Exp((CARRY_COST - RATE) * SECOND_TENOR) * _
          (CBND_FUNC(D1_VAL, E1_VAL, RHO_VAL, CND_TYPE, CBND_TYPE) - (TEMP_BARRIER / SPOT) ^ (2 * _
          (MU_VAL + 1)) * CBND_FUNC(F1_VAL, -E3_VAL, -RHO_VAL, CND_TYPE, CBND_TYPE)) _
            - STRIKE * Exp(-RATE * SECOND_TENOR) * (CBND_FUNC(D2_VAL, E2_VAL, RHO_VAL, CND_TYPE, _
            CBND_TYPE) - (TEMP_BARRIER / SPOT) ^ _
            (2 * MU_VAL) * CBND_FUNC(F2_VAL, -E4_VAL, -RHO_VAL, CND_TYPE, CBND_TYPE))
    
    ElseIf ((OPTION_FLAG = 5)) And (STRIKE < TEMP_BARRIER) Then
        '(OPTION_FLAG = "cob1") Or
        PARTIAL_TIME_SINGLE_ASSET_BARRIER_OPTION_FUNC = SPOT * Exp((CARRY_COST - RATE) * SECOND_TENOR) * _
             (CBND_FUNC(-G1_VAL, -E1_VAL, RHO_VAL, CND_TYPE, CBND_TYPE) - (TEMP_BARRIER / SPOT) ^ _
             (2 * (MU_VAL + 1)) * CBND_FUNC(-G3_VAL, E3_VAL, -RHO_VAL, CND_TYPE, CBND_TYPE)) _
            - STRIKE * Exp(-RATE * SECOND_TENOR) * (CBND_FUNC(-G2_VAL, -E2_VAL, RHO_VAL, CND_TYPE, _
            CBND_TYPE) - (TEMP_BARRIER / SPOT) ^ _
            (2 * MU_VAL) * CBND_FUNC(-G4_VAL, E4_VAL, -RHO_VAL, CND_TYPE, CBND_TYPE)) - SPOT * _
            Exp((CARRY_COST - RATE) * SECOND_TENOR) _
            * (CBND_FUNC(-D1_VAL, -E1_VAL, RHO_VAL, CND_TYPE, CBND_TYPE) - (TEMP_BARRIER / _
            SPOT) ^ (2 * (MU_VAL + 1)) * CBND_FUNC(-F1_VAL, E3_VAL, -RHO_VAL, CND_TYPE, CBND_TYPE)) _
            + STRIKE * Exp(-RATE * SECOND_TENOR) * (CBND_FUNC(-D2_VAL, -E2_VAL, RHO_VAL, _
            CND_TYPE, CBND_TYPE) - (TEMP_BARRIER / SPOT) ^ _
            (2 * MU_VAL) * CBND_FUNC(-F2_VAL, E4_VAL, -RHO_VAL, CND_TYPE, CBND_TYPE)) + SPOT * _
            Exp((CARRY_COST - RATE) * SECOND_TENOR) _
            * (CBND_FUNC(G1_VAL, E1_VAL, RHO_VAL, CND_TYPE, CBND_TYPE) - (TEMP_BARRIER / SPOT) ^ _
            (2 * (MU_VAL + 1)) * CBND_FUNC(G3_VAL, -E3_VAL, -RHO_VAL, CND_TYPE, CBND_TYPE)) _
            - STRIKE * Exp(-RATE * SECOND_TENOR) * (CBND_FUNC(G2_VAL, E2_VAL, RHO_VAL, CND_TYPE, _
            CBND_TYPE) - (TEMP_BARRIER / SPOT) ^ _
            (2 * MU_VAL) * CBND_FUNC(G4_VAL, -E4_VAL, -RHO_VAL, CND_TYPE, CBND_TYPE))

    '---------------------------------------------------------------------------
    '------------------------put down-and out and up-and-out type a
    '---------------------------------------------------------------------------

    ElseIf (OPTION_FLAG = 4) Then
        '(OPTION_FLAG = "pdoA") Or
        PARTIAL_TIME_SINGLE_ASSET_BARRIER_OPTION_FUNC = _
        PARTIAL_TIME_SINGLE_ASSET_BARRIER_OPTION_FUNC(SPOT, STRIKE, BARRIER, _
            FIRST_TENOR, SECOND_TENOR, RATE, CARRY_COST, SIGMA, 2, MONITOR_BASIS, _
            ADJ_FACTOR) - SPOT * Exp((CARRY_COST - RATE) * SECOND_TENOR) * Z5_VAL + _
            STRIKE * Exp(-RATE * SECOND_TENOR) * Z1_VAL
    
    ElseIf (OPTION_FLAG = 3) Then
        'OPTION_FLAG = "puoA" Or
        PARTIAL_TIME_SINGLE_ASSET_BARRIER_OPTION_FUNC = _
        PARTIAL_TIME_SINGLE_ASSET_BARRIER_OPTION_FUNC(SPOT, STRIKE, BARRIER, _
            FIRST_TENOR, SECOND_TENOR, RATE, CARRY_COST, SIGMA, 1, MONITOR_BASIS, _
            ADJ_FACTOR) - SPOT * Exp((CARRY_COST - _
            RATE) * SECOND_TENOR) * Z6_VAL + STRIKE * Exp(-RATE * SECOND_TENOR) * Z2_VAL

    '---------------------------------------------------------------------------
    '------------------------------put out type a1
    '---------------------------------------------------------------------------
    
    ElseIf (OPTION_FLAG = 6) Then
        '(OPTION_FLAG = "pob1") Or
        PARTIAL_TIME_SINGLE_ASSET_BARRIER_OPTION_FUNC = _
        PARTIAL_TIME_SINGLE_ASSET_BARRIER_OPTION_FUNC(SPOT, STRIKE, BARRIER, _
            FIRST_TENOR, SECOND_TENOR, RATE, CARRY_COST, SIGMA, 5, MONITOR_BASIS, _
            ADJ_FACTOR) _
            - SPOT * Exp((CARRY_COST - RATE) * SECOND_TENOR) * Z8_VAL + STRIKE * _
            Exp(-RATE * SECOND_TENOR) * Z4_VAL - SPOT * Exp((CARRY_COST - RATE) * _
            SECOND_TENOR) * Z7_VAL + STRIKE * Exp(-RATE * SECOND_TENOR) * Z3_VAL

    '---------------------------------------------------------------------------
    '------------------------------put down-and-out type b1
    '---------------------------------------------------------------------------

    ElseIf (OPTION_FLAG = 10) Then
        '(OPTION_FLAG = "pdob2") Or
        PARTIAL_TIME_SINGLE_ASSET_BARRIER_OPTION_FUNC = _
        PARTIAL_TIME_SINGLE_ASSET_BARRIER_OPTION_FUNC(SPOT, STRIKE, BARRIER, _
            FIRST_TENOR, SECOND_TENOR, RATE, CARRY_COST, SIGMA, 8, MONITOR_BASIS, _
            ADJ_FACTOR) _
            - SPOT * Exp((CARRY_COST - RATE) * SECOND_TENOR) * Z7_VAL + STRIKE * _
            Exp(-RATE * SECOND_TENOR) * Z3_VAL
    
    '---------------------------------------------------------------------------
    '------------------------------put up-and-out type b1
    '---------------------------------------------------------------------------
    
    Else 'If (OPTION_FLAG = "puob2") Or (OPTION_FLAG = 9) Then
        PARTIAL_TIME_SINGLE_ASSET_BARRIER_OPTION_FUNC = _
        PARTIAL_TIME_SINGLE_ASSET_BARRIER_OPTION_FUNC(SPOT, STRIKE, BARRIER, _
            FIRST_TENOR, SECOND_TENOR, RATE, CARRY_COST, SIGMA, 7, MONITOR_BASIS, _
            ADJ_FACTOR) _
            - SPOT * Exp((CARRY_COST - RATE) * SECOND_TENOR) * Z8_VAL + STRIKE * _
            Exp(-RATE * SECOND_TENOR) * Z4_VAL
        
    End If

Exit Function
ERROR_LABEL:
PARTIAL_TIME_SINGLE_ASSET_BARRIER_OPTION_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : PARTIAL_TIME_TWO_ASSET_BARRIER_OPTION_FUNC
'DESCRIPTION   : Partial-time two asset barrier options
'LIBRARY       : DERIVATIVES
'GROUP         : BARRIER
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009

'************************************************************************************
'************************************************************************************

Function PARTIAL_TIME_TWO_ASSET_BARRIER_OPTION_FUNC(ByVal SPOT_A As Double, _
ByVal SPOT_B As Double, _
ByVal STRIKE As Double, _
ByVal BARRIER As Double, _
ByVal FIRST_TENOR As Double, _
ByVal SECOND_TENOR As Double, _
ByVal RATE As Double, _
ByVal CARRY_COST_A As Double, _
ByVal CARRY_COST_B As Double, _
ByVal SIGMA_A As Double, _
ByVal SIGMA_B As Double, _
ByVal RHO_VAL As Double, _
Optional ByVal OPTION_FLAG As Integer = 1, _
Optional ByVal MONITOR_BASIS As Integer = 2, _
Optional ByVal ADJ_FACTOR As Double = 0.5826, _
Optional ByVal CND_TYPE As Integer = 0, _
Optional ByVal CBND_TYPE As Integer = 0)


'OPTION_FLAG Options
    
    '1: Down-and-in call --> cdi
    '2: Up-and-in call --> cui
    '3: Down-and-in put --> pdi
    '4: Up-and-in put --> pui
    '5: Down-and-out call --> cdo
    '6: Up-and-out call --> cuo
    '7: Down-and-out put --> pdo
    '8: Up-and-out put --> puo


    Dim D1_VAL As Double
    Dim D2_VAL As Double
    Dim D3_VAL As Double
    Dim D4_VAL As Double
    
    Dim E1_VAL As Double
    Dim E2_VAL As Double
    
    Dim E3_VAL As Double
    Dim E4_VAL As Double
    
    Dim BMU_VAL As Double
    Dim AMU_VAL As Double

    Dim ETA_VAL As Integer
    Dim PHI_VAL As Integer
    
    Dim TEMP_VAL As Double
    
    On Error GoTo ERROR_LABEL
    
    BARRIER = DISCRETE_BARRIER_MONITORING_ADJ_FUNC(SPOT_B, BARRIER, SIGMA_B, _
    BARRIER_MONITORING_COUNT_BASIS_FUNC(MONITOR_BASIS), ADJ_FACTOR)

    Select Case OPTION_FLAG
        Case 5, 7, 1, 3 '"cdo", "pdo", "cdi", "pdi"
            PHI_VAL = -1
        Case Else
            PHI_VAL = 1
    End Select
    
    Select Case OPTION_FLAG
        Case 5, 6, 1, 2 '"cdo", "cuo", "cdi", "cui"
            ETA_VAL = 1
        Case Else
            ETA_VAL = -1
    End Select
    
    BMU_VAL = CARRY_COST_A - SIGMA_A ^ 2 / 2
    AMU_VAL = CARRY_COST_B - SIGMA_B ^ 2 / 2
    
    D1_VAL = (Log(SPOT_A / STRIKE) + (BMU_VAL + SIGMA_A ^ 2) * SECOND_TENOR) / _
    (SIGMA_A * Sqr(SECOND_TENOR))
    
    D2_VAL = D1_VAL - SIGMA_A * Sqr(SECOND_TENOR)
    
    D3_VAL = D1_VAL + 2 * RHO_VAL * Log(BARRIER / SPOT_B) / (SIGMA_B * Sqr(SECOND_TENOR))
    
    D4_VAL = D2_VAL + 2 * RHO_VAL * Log(BARRIER / SPOT_B) / (SIGMA_B * Sqr(SECOND_TENOR))
    
    E1_VAL = (Log(BARRIER / SPOT_B) - (AMU_VAL + RHO_VAL * SIGMA_A * SIGMA_B) * FIRST_TENOR) / _
    (SIGMA_B * Sqr(FIRST_TENOR))
    
    E2_VAL = E1_VAL + RHO_VAL * SIGMA_A * Sqr(FIRST_TENOR)
    E3_VAL = E1_VAL - 2 * Log(BARRIER / SPOT_B) / (SIGMA_B * Sqr(FIRST_TENOR))
    E4_VAL = E2_VAL - 2 * Log(BARRIER / SPOT_B) / (SIGMA_B * Sqr(FIRST_TENOR))

    TEMP_VAL = ETA_VAL * SPOT_A * Exp((CARRY_COST_A - RATE) * SECOND_TENOR) * _
        (CBND_FUNC(ETA_VAL * D1_VAL, PHI_VAL * E1_VAL, -ETA_VAL * PHI_VAL * RHO_VAL * Sqr(FIRST_TENOR / _
        SECOND_TENOR), CND_TYPE, CBND_TYPE) _
        - Exp(2 * Log(BARRIER / SPOT_B) * (AMU_VAL + RHO_VAL * SIGMA_A * SIGMA_B) / _
        (SIGMA_B ^ 2)) _
        * CBND_FUNC(ETA_VAL * D3_VAL, PHI_VAL * E3_VAL, -ETA_VAL * PHI_VAL * RHO_VAL * Sqr(FIRST_TENOR / _
        SECOND_TENOR), CND_TYPE, CBND_TYPE)) _
        - ETA_VAL * Exp(-RATE * SECOND_TENOR) * STRIKE * (CBND_FUNC(ETA_VAL * D2_VAL, PHI_VAL * _
        E2_VAL, -ETA_VAL * PHI_VAL _
        * RHO_VAL * Sqr(FIRST_TENOR / SECOND_TENOR), CND_TYPE, CBND_TYPE) - _
        Exp(2 * Log(BARRIER / SPOT_B) * AMU_VAL / _
        (SIGMA_B ^ 2)) * CBND_FUNC(ETA_VAL * D4_VAL, PHI_VAL * E4_VAL, -ETA_VAL * PHI_VAL * RHO_VAL * _
        Sqr(FIRST_TENOR / _
        SECOND_TENOR), CND_TYPE, CBND_TYPE))
    
    Select Case OPTION_FLAG
        Case 5, 6, 7, 8 '"cdo", "cuo", "pdo", "puo"
            PARTIAL_TIME_TWO_ASSET_BARRIER_OPTION_FUNC = TEMP_VAL
        Case 2, 1 '"cui", "cdi"
            PARTIAL_TIME_TWO_ASSET_BARRIER_OPTION_FUNC = GENERALIZED_BLACK_SCHOLES_FUNC(SPOT_A, STRIKE, _
            SECOND_TENOR, RATE, CARRY_COST_A, SIGMA_A, 1, CND_TYPE) - TEMP_VAL
        Case Else '4, 3 ' "pui", "pdi"
            PARTIAL_TIME_TWO_ASSET_BARRIER_OPTION_FUNC = GENERALIZED_BLACK_SCHOLES_FUNC(SPOT_A, STRIKE, _
            SECOND_TENOR, RATE, CARRY_COST_A, SIGMA_A, -1, CND_TYPE) - TEMP_VAL
    End Select
    
Exit Function
ERROR_LABEL:
PARTIAL_TIME_TWO_ASSET_BARRIER_OPTION_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : LOOK_BARRIER_OPTION_FUNC
'DESCRIPTION   : Look-barrier options
'LIBRARY       : DERIVATIVES
'GROUP         : BARRIER
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009

'************************************************************************************
'************************************************************************************

Function LOOK_BARRIER_OPTION_FUNC(ByVal SPOT As Double, _
ByVal STRIKE As Double, _
ByVal BARRIER As Double, _
ByVal FIRST_TENOR As Double, _
ByVal SECOND_TENOR As Double, _
ByVal RATE As Double, _
ByVal CARRY_COST As Double, _
ByVal SIGMA As Double, _
Optional ByVal OPTION_FLAG As Integer = 1, _
Optional ByVal MONITOR_BASIS As Integer = 1, _
Optional ByVal ADJ_FACTOR As Double = 0.5826, _
Optional ByVal CND_TYPE As Integer = 0, _
Optional ByVal CBND_TYPE As Integer = 0)
    
    Dim BMU_VAL As Double
    Dim AMU_VAL As Double
    
    Dim RHO_VAL As Double
    Dim MIN_MAX_VAL As Double
    
    Dim G1_VAL As Double
    Dim G2_VAL As Double
    
    Dim ATEMP_VAL As Double
    Dim BTEMP_VAL As Double
    Dim CTEMP_VAL As Double
    Dim DTEMP_VAL As Double
    
    Dim ETA_VAL As Double
    
    Dim TEMP_CASH As Double
    Dim TEMP_BARRIER As Double
    
    On Error GoTo ERROR_LABEL
    
    '1 Up-and-out call (BARRIER>SPOT): cuo
    '2 Up-and-in call (BARRIER>SPOT): cui
    '3 Down-and-out put (BARRIER<SPOT): pdo
    '4 Down-and-in put (BARRIER<SPOT): pdi
    
    BARRIER = DISCRETE_BARRIER_MONITORING_ADJ_FUNC(SPOT, BARRIER, SIGMA, _
    BARRIER_MONITORING_COUNT_BASIS_FUNC(MONITOR_BASIS), ADJ_FACTOR)
    
    TEMP_BARRIER = Log(BARRIER / SPOT)
    TEMP_CASH = Log(STRIKE / SPOT)
    BMU_VAL = CARRY_COST - SIGMA ^ 2 / 2
    AMU_VAL = CARRY_COST + SIGMA ^ 2 / 2
    RHO_VAL = Sqr(FIRST_TENOR / SECOND_TENOR)
    
        Select Case OPTION_FLAG
            Case 1, 2 ', "cuo", "cui"
                ETA_VAL = 1
                MIN_MAX_VAL = MINIMUM_FUNC(TEMP_BARRIER, TEMP_CASH)
            Case Else '3, 4 ' "pdo", "pdi"
                ETA_VAL = -1
                MIN_MAX_VAL = MAXIMUM_FUNC(TEMP_BARRIER, TEMP_CASH)
        End Select
    
    G1_VAL = (CND_FUNC(ETA_VAL * (TEMP_BARRIER - AMU_VAL * FIRST_TENOR) / (SIGMA * _
        Sqr(FIRST_TENOR)), CND_TYPE) - _
       Exp(2 * AMU_VAL * TEMP_BARRIER / SIGMA ^ 2) * CND_FUNC(ETA_VAL * _
       (-TEMP_BARRIER - AMU_VAL * _
        FIRST_TENOR) / (SIGMA * Sqr(FIRST_TENOR)), CND_TYPE)) - _
        (CND_FUNC(ETA_VAL * (MIN_MAX_VAL - AMU_VAL * _
        FIRST_TENOR) / (SIGMA * Sqr(FIRST_TENOR)), CND_TYPE) - _
        Exp(2 * AMU_VAL * TEMP_BARRIER / _
        SIGMA ^ 2) * CND_FUNC(ETA_VAL * (MIN_MAX_VAL - 2 * TEMP_BARRIER - _
        AMU_VAL * FIRST_TENOR) / _
        (SIGMA * Sqr(FIRST_TENOR)), CND_TYPE))
    
    G2_VAL = (CND_FUNC(ETA_VAL * (TEMP_BARRIER - BMU_VAL * FIRST_TENOR) / _
         (SIGMA * Sqr(FIRST_TENOR)), CND_TYPE) - _
        Exp(2 * BMU_VAL * TEMP_BARRIER / SIGMA ^ 2) * CND_FUNC(ETA_VAL * _
        (-TEMP_BARRIER - BMU_VAL * _
        FIRST_TENOR) / (SIGMA * Sqr(FIRST_TENOR)), CND_TYPE)) - _
        (CND_FUNC(ETA_VAL * _
        (MIN_MAX_VAL - BMU_VAL * FIRST_TENOR) / (SIGMA * Sqr(FIRST_TENOR)), _
        CND_TYPE) - _
        Exp(2 * BMU_VAL * TEMP_BARRIER / SIGMA ^ 2) * CND_FUNC(ETA_VAL * _
        (MIN_MAX_VAL - 2 * _
        TEMP_BARRIER - BMU_VAL * FIRST_TENOR) / (SIGMA * _
        Sqr(FIRST_TENOR)), CND_TYPE))

    ATEMP_VAL = SPOT * Exp((CARRY_COST - RATE) * SECOND_TENOR) * (1 + SIGMA ^ 2 / _
        (2 * CARRY_COST)) * (CBND_FUNC(ETA_VAL * (MIN_MAX_VAL - AMU_VAL * FIRST_TENOR) / (SIGMA * _
        Sqr(FIRST_TENOR)), ETA_VAL * (-TEMP_CASH + AMU_VAL * SECOND_TENOR) / (SIGMA * _
        Sqr(SECOND_TENOR)), -RHO_VAL, CND_TYPE, CBND_TYPE) - Exp(2 * AMU_VAL * _
        TEMP_BARRIER / SIGMA ^ 2) _
        * CBND_FUNC(ETA_VAL * (MIN_MAX_VAL - 2 * TEMP_BARRIER - AMU_VAL * FIRST_TENOR) / (SIGMA * _
        Sqr(FIRST_TENOR)), ETA_VAL * (2 * TEMP_BARRIER - TEMP_CASH + AMU_VAL * _
        SECOND_TENOR) / _
        (SIGMA * Sqr(SECOND_TENOR)), -RHO_VAL, CND_TYPE, CBND_TYPE))
    
    BTEMP_VAL = -Exp(-RATE * SECOND_TENOR) * STRIKE * (CBND_FUNC(ETA_VAL * (MIN_MAX_VAL - BMU_VAL * _
       FIRST_TENOR) / (SIGMA * Sqr(FIRST_TENOR)), ETA_VAL * (-TEMP_CASH + BMU_VAL * _
        SECOND_TENOR) / (SIGMA * Sqr(SECOND_TENOR)), -RHO_VAL, CND_TYPE, _
        CBND_TYPE) - Exp(2 * BMU_VAL * TEMP_BARRIER / _
        SIGMA ^ 2) * CBND_FUNC(ETA_VAL * (MIN_MAX_VAL - 2 * TEMP_BARRIER - BMU_VAL * _
        FIRST_TENOR) / _
        (SIGMA * Sqr(FIRST_TENOR)), ETA_VAL * (2 * TEMP_BARRIER - TEMP_CASH + _
        BMU_VAL * SECOND_TENOR) / (SIGMA * Sqr(SECOND_TENOR)), -RHO_VAL, _
        CND_TYPE, CBND_TYPE))
    
    CTEMP_VAL = -Exp(-RATE * SECOND_TENOR) * SIGMA ^ 2 / (2 * CARRY_COST) * (SPOT * _
        (SPOT / STRIKE) ^ (-2 * CARRY_COST / SIGMA ^ 2) * CBND_FUNC(ETA_VAL * (MIN_MAX_VAL + _
        BMU_VAL * _
        FIRST_TENOR) / (SIGMA * Sqr(FIRST_TENOR)), ETA_VAL * (-TEMP_CASH - BMU_VAL * _
        SECOND_TENOR) / _
        (SIGMA * Sqr(SECOND_TENOR)), -RHO_VAL, CND_TYPE, CBND_TYPE) - BARRIER * _
        (BARRIER / STRIKE) ^ (-2 * _
        CARRY_COST / SIGMA ^ 2) * CBND_FUNC(ETA_VAL * (MIN_MAX_VAL - 2 * TEMP_BARRIER + BMU_VAL * _
        FIRST_TENOR) / (SIGMA * Sqr(FIRST_TENOR)), ETA_VAL * (2 * TEMP_BARRIER - _
        TEMP_CASH - _
        BMU_VAL * SECOND_TENOR) / (SIGMA * Sqr(SECOND_TENOR)), -RHO_VAL, CND_TYPE, _
        CBND_TYPE))
    
    DTEMP_VAL = SPOT * Exp((CARRY_COST - RATE) * SECOND_TENOR) * ((1 + SIGMA ^ 2 / _
        (2 * CARRY_COST)) * CND_FUNC(ETA_VAL * AMU_VAL * (SECOND_TENOR - FIRST_TENOR) / _
        (SIGMA * _
        Sqr(SECOND_TENOR - FIRST_TENOR)), CND_TYPE) + Exp(-CARRY_COST * _
        (SECOND_TENOR - FIRST_TENOR)) _
        * (1 - SIGMA ^ 2 / (2 * CARRY_COST)) * CND_FUNC(ETA_VAL * (-BMU_VAL * _
        (SECOND_TENOR - _
        FIRST_TENOR)) / (SIGMA * Sqr(SECOND_TENOR - FIRST_TENOR)), CND_TYPE)) * _
        G1_VAL - Exp(-RATE * SECOND_TENOR) * STRIKE * G2_VAL

        Select Case OPTION_FLAG
            Case 1, 3 ', "cuo", "pdo"
                LOOK_BARRIER_OPTION_FUNC = ETA_VAL * (ATEMP_VAL + BTEMP_VAL + CTEMP_VAL + DTEMP_VAL)
            Case 2 ', "cui"
                LOOK_BARRIER_OPTION_FUNC = PARTIAL_TIME_FIXED_STRIKE_LOOKBACK_OPTION_FUNC(SPOT, STRIKE, FIRST_TENOR, _
                SECOND_TENOR, RATE, CARRY_COST, SIGMA, 1) - (ETA_VAL * (ATEMP_VAL + BTEMP_VAL + _
                CTEMP_VAL + DTEMP_VAL))
            Case Else '4, "pdi"
                LOOK_BARRIER_OPTION_FUNC = PARTIAL_TIME_FIXED_STRIKE_LOOKBACK_OPTION_FUNC(SPOT, STRIKE, FIRST_TENOR, _
                SECOND_TENOR, RATE, CARRY_COST, SIGMA, -1) - (ETA_VAL * (ATEMP_VAL + BTEMP_VAL + _
                CTEMP_VAL + DTEMP_VAL))
        End Select
    
Exit Function
ERROR_LABEL:
LOOK_BARRIER_OPTION_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : SOFT_BARRIER_OPTION_FUNC
'DESCRIPTION   : Soft barrier options
'LIBRARY       : DERIVATIVES
'GROUP         : BARRIER
'ID            : 007
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009

'************************************************************************************
'************************************************************************************

Function SOFT_BARRIER_OPTION_FUNC(ByVal SPOT As Double, _
ByVal cash As Double, _
ByVal LOWER_BARRIER As Double, _
ByVal UPPER_BARRIER As Double, _
ByVal TENOR As Double, _
ByVal RATE As Double, _
ByVal CARRY_COST As Double, _
ByVal SIGMA As Double, _
Optional ByVal OPTION_FLAG As Integer = 1, _
Optional ByVal CND_TYPE As Integer = 0)

    Dim MU_VAL As Double
    
    Dim D1_VAL As Double
    Dim D2_VAL As Double
    Dim D3_VAL As Double
    Dim D4_VAL As Double
    
    Dim E1_VAL As Double
    Dim E2_VAL As Double
    Dim E3_VAL As Double
    Dim E4_VAL As Double
    
    Dim ALAMBDA_VAL As Double
    Dim BLAMBDA_VAL As Double
    
    Dim TEMP_VAL As Double
    Dim ETA_VAL As Integer
    
    On Error GoTo ERROR_LABEL

   'OPTION_FLAG: [1] Down-and-in call
   'OPTION_FLAG: [2] Down-and-out call
   'OPTION_FLAG: [3] Up-and-in put
   'OPTION_FLAG: [4] Up-and-out put
   
   Select Case OPTION_FLAG
        Case 1, 2 ', "cdi", "cdo"
            ETA_VAL = 1
        Case Else '3, 4, "pui", "puo"
            ETA_VAL = -1
   End Select
    
    MU_VAL = (CARRY_COST + SIGMA ^ 2 / 2) / SIGMA ^ 2
    ALAMBDA_VAL = Exp(-1 / 2 * SIGMA ^ 2 * TENOR * (MU_VAL + 0.5) * (MU_VAL - 0.5))
    BLAMBDA_VAL = Exp(-1 / 2 * SIGMA ^ 2 * TENOR * (MU_VAL - 0.5) * (MU_VAL - 1.5))
    
    D1_VAL = Log(UPPER_BARRIER ^ 2 / (SPOT * cash)) / (SIGMA * Sqr(TENOR)) + MU_VAL * _
    SIGMA * Sqr(TENOR)
    D2_VAL = D1_VAL - (MU_VAL + 0.5) * SIGMA * Sqr(TENOR)
    
    D3_VAL = Log(UPPER_BARRIER ^ 2 / (SPOT * cash)) / (SIGMA * Sqr(TENOR)) + (MU_VAL - 1) * _
    SIGMA * Sqr(TENOR)
    D4_VAL = D3_VAL - (MU_VAL - 0.5) * SIGMA * Sqr(TENOR)
    
    E1_VAL = Log(LOWER_BARRIER ^ 2 / (SPOT * cash)) / (SIGMA * Sqr(TENOR)) + MU_VAL * SIGMA * _
    Sqr(TENOR)
    E2_VAL = E1_VAL - (MU_VAL + 0.5) * SIGMA * Sqr(TENOR)
    
    E3_VAL = Log(LOWER_BARRIER ^ 2 / (SPOT * cash)) / (SIGMA * Sqr(TENOR)) + (MU_VAL - 1) * _
    SIGMA * Sqr(TENOR)
    E4_VAL = E3_VAL - (MU_VAL - 0.5) * SIGMA * Sqr(TENOR)
    
    TEMP_VAL = ETA_VAL * 1 / (UPPER_BARRIER - LOWER_BARRIER) * (SPOT * Exp((CARRY_COST - RATE) * _
        TENOR) * SPOT ^ (-2 * MU_VAL) * (SPOT * cash) ^ (MU_VAL + 0.5) / (2 * (MU_VAL + 0.5)) _
        * ((UPPER_BARRIER ^ 2 / (SPOT * cash)) ^ (MU_VAL + 0.5) * CND_FUNC(ETA_VAL * D1_VAL, _
        CND_TYPE) - ALAMBDA_VAL * _
        CND_FUNC(ETA_VAL * D2_VAL, CND_TYPE) - (LOWER_BARRIER ^ 2 / (SPOT * cash)) ^ _
        (MU_VAL + 0.5) * CND_FUNC(ETA_VAL * E1_VAL, CND_TYPE) + _
        ALAMBDA_VAL * CND_FUNC(ETA_VAL * E2_VAL, CND_TYPE)) - cash * Exp(-RATE * TENOR) * _
        SPOT ^ (-2 * (MU_VAL - 1)) _
        * (SPOT * cash) ^ (MU_VAL - 0.5) / (2 * (MU_VAL - 0.5)) * ((UPPER_BARRIER ^ 2 _
        / (SPOT * _
        cash)) ^ (MU_VAL - 0.5) * CND_FUNC(ETA_VAL * D3_VAL, CND_TYPE) - BLAMBDA_VAL * CND_FUNC(ETA_VAL * _
        D4_VAL, CND_TYPE) _
        - (LOWER_BARRIER ^ 2 / (SPOT * cash)) ^ (MU_VAL - 0.5) * CND_FUNC(ETA_VAL * E3_VAL, _
        CND_TYPE) + BLAMBDA_VAL * _
        CND_FUNC(ETA_VAL * E4_VAL, CND_TYPE)))
    

   Select Case OPTION_FLAG
        Case 1, 3 ', "cdi", "pui"
            SOFT_BARRIER_OPTION_FUNC = TEMP_VAL
        Case 2 ', "cdo"
            SOFT_BARRIER_OPTION_FUNC = GENERALIZED_BLACK_SCHOLES_FUNC(SPOT, cash, TENOR, _
            RATE, CARRY_COST, SIGMA, 1, CND_TYPE) - TEMP_VAL
        Case Else '4 ', "puo"
            SOFT_BARRIER_OPTION_FUNC = GENERALIZED_BLACK_SCHOLES_FUNC(SPOT, cash, TENOR, RATE, _
            CARRY_COST, SIGMA, -1, CND_TYPE) - TEMP_VAL
   End Select

Exit Function
ERROR_LABEL:
SOFT_BARRIER_OPTION_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : BINARY_BARRIER_OPTION_FUNC

'DESCRIPTION   : Binary barrier options
'Binary options behave similarly to standard options, but the payout
'is based on whether the option is on the money, not by how much it is
'in the money. For this reason they are also called all-or-nothing options.

'As with a standard European style option, the payoff is based on the price
'of the underlying asset on the expiration date. Unlike with standard options,
'the payoff is fixed at the writing of the contract. Binary options may
'incorporate barrier features, which are as described in the
'Barrier Functions section:

'1) Down and Out
    'The option is canceled or knocked-out if the asset falls to a
    'predetermined boundary price.
'2) Down and In
    'The option is activated or knocked-in if the asset falls to a
    'predetermined boundary price.
'3) Up and Out
    'The option is canceled or knocked-out if the asset rises to a
    'predetermined boundary price.
'4) Up and In
    'The option is activated or knocked-in if the asset rises to a
    'predetermined boundary price.
'Benefits

'The purchaser and writer of binary options need only determine an expected
'direction of price movement, rather the direction and the magnitude, in order
'to effectively use the option.

'Features
'The payout profile (and hence sensitivity to price changes) of binary options
'is discontinuous.

'The cash payoff may be equal to the strike price, or it may be greater or less,
'in which case the option is a gap option.

'Intrinsic Value Formula
'The payoff at expiration is:
'Call: 0, if U < k; X, if U > k
'Put: 0, if U > k, X, if U < k

'Where U is the underlying price, E is the exercise or strike price, and X is
'the payoff for a cash-or-nothing, or U, for an asset-or-nothing option;
'additional condition is that D&O and U&O binary barriers are worth R or
'rebate if barrier is reached.

'Aliases
'Binary options are also known as digital options and all-or-nothing options.
'Uses

'1) A bank wishes to hedge a key interest rate exceeding a certain level. It
'purchases a binary option with strike at the level at which they wish to
'hedge. If the interest rate exceeds that level, they receive a fixed payment,
'no matter how high the rate goes.

'LIBRARY       : DERIVATIVES
'GROUP         : BARRIER
'ID            : 008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009

'************************************************************************************
'************************************************************************************

Function BINARY_BARRIER_OPTION_FUNC(ByVal SPOT As Double, _
ByVal STRIKE As Double, _
ByVal BARRIER As Double, _
ByVal cash As Double, _
ByVal TENOR As Double, _
ByVal RATE As Double, _
ByVal CARRY_COST As Double, _
ByVal SIGMA As Double, _
Optional ByVal OPTION_FLAG As Integer = 1, _
Optional ByVal MONITOR_BASIS As Integer = 2, _
Optional ByVal ADJ_FACTOR As Double = 0.5826, _
Optional ByVal CND_TYPE As Integer = 0)

    '// OPTION_FLAG:  Value 1 to 28 dependent on binary option type
    
    Dim X1_VAL As Double
    Dim X2_VAL As Double
    
    Dim Y1_VAL As Double
    Dim Y2_VAL As Double
    
    Dim Z_VAL As Double
    Dim MU_VAL As Double
    Dim LAMBDA As Double
    
    Dim A1_VAL As Double
    Dim A2_VAL As Double
    Dim A3_VAL As Double
    Dim A4_VAL As Double
    Dim A5_VAL As Double
    
    Dim B1_VAL As Double
    Dim B2_VAL As Double
    Dim B3_VAL As Double
    Dim B4_VAL As Double

    Dim PHI_VAL As Integer
    Dim ETA_VAL As Integer
    
    On Error GoTo ERROR_LABEL

    PHI_VAL = BARRIER_BINARY_PHI_FACTOR_FUNC(OPTION_FLAG)
    ETA_VAL = BARRIER_BINARY_ETA_FACTOR_FUNC(OPTION_FLAG)
    
    BARRIER = DISCRETE_BARRIER_MONITORING_ADJ_FUNC(SPOT, BARRIER, SIGMA, _
    BARRIER_MONITORING_COUNT_BASIS_FUNC(MONITOR_BASIS), ADJ_FACTOR)
    
    MU_VAL = (CARRY_COST - SIGMA ^ 2 / 2) / SIGMA ^ 2
    
    LAMBDA = Sqr(MU_VAL ^ 2 + 2 * RATE / SIGMA ^ 2)
    
    X1_VAL = Log(SPOT / STRIKE) / (SIGMA * Sqr(TENOR)) + (MU_VAL + 1) * SIGMA * Sqr(TENOR)
    
    X2_VAL = Log(SPOT / BARRIER) / (SIGMA * Sqr(TENOR)) + (MU_VAL + 1) * SIGMA * Sqr(TENOR)
    
    Y1_VAL = Log(BARRIER ^ 2 / (SPOT * STRIKE)) / (SIGMA * Sqr(TENOR)) + (MU_VAL + 1) * _
    SIGMA * Sqr(TENOR)
    
    Y2_VAL = Log(BARRIER / SPOT) / (SIGMA * Sqr(TENOR)) + (MU_VAL + 1) * SIGMA * Sqr(TENOR)
    
    Z_VAL = Log(BARRIER / SPOT) / (SIGMA * Sqr(TENOR)) + LAMBDA * SIGMA * Sqr(TENOR)
    
    A1_VAL = SPOT * Exp((CARRY_COST - RATE) * TENOR) * CND_FUNC(PHI_VAL * X1_VAL, CND_TYPE)
    B1_VAL = cash * Exp(-RATE * TENOR) * CND_FUNC(PHI_VAL * X1_VAL - PHI_VAL * SIGMA * Sqr(TENOR), _
        CND_TYPE)
    
    A2_VAL = SPOT * Exp((CARRY_COST - RATE) * TENOR) * CND_FUNC(PHI_VAL * X2_VAL, CND_TYPE)
    B2_VAL = cash * Exp(-RATE * TENOR) * CND_FUNC(PHI_VAL * X2_VAL - PHI_VAL * SIGMA * Sqr(TENOR), _
        CND_TYPE)
    
    A3_VAL = SPOT * Exp((CARRY_COST - RATE) * TENOR) * (BARRIER / SPOT) ^ (2 * _
        (MU_VAL + 1)) * CND_FUNC(ETA_VAL * Y1_VAL, CND_TYPE)
    
    B3_VAL = cash * Exp(-RATE * TENOR) * (BARRIER / SPOT) ^ (2 * MU_VAL) * CND_FUNC(ETA_VAL * _
        Y1_VAL - ETA_VAL * SIGMA * Sqr(TENOR), CND_TYPE)
    
    A4_VAL = SPOT * Exp((CARRY_COST - RATE) * TENOR) * (BARRIER / SPOT) ^ (2 * _
        (MU_VAL + 1)) * CND_FUNC(ETA_VAL * Y2_VAL, CND_TYPE)
    
    B4_VAL = cash * Exp(-RATE * TENOR) * (BARRIER / SPOT) ^ (2 * MU_VAL) * CND_FUNC(ETA_VAL * _
        Y2_VAL - ETA_VAL * SIGMA * Sqr(TENOR), CND_TYPE)
    
    A5_VAL = cash * ((BARRIER / SPOT) ^ (MU_VAL + LAMBDA) * CND_FUNC(ETA_VAL * Z_VAL, CND_TYPE) + _
        (BARRIER / SPOT) ^ (MU_VAL - LAMBDA) * CND_FUNC(ETA_VAL * Z_VAL - 2 * ETA_VAL * LAMBDA * _
        SIGMA * Sqr(TENOR), CND_TYPE))
    
    If STRIKE > BARRIER Then
        Select Case OPTION_FLAG
            Case Is < 5
                BINARY_BARRIER_OPTION_FUNC = A5_VAL
            Case Is < 7
                BINARY_BARRIER_OPTION_FUNC = B2_VAL + B4_VAL
            Case Is < 9
                BINARY_BARRIER_OPTION_FUNC = A2_VAL + A4_VAL
            Case Is < 11
                BINARY_BARRIER_OPTION_FUNC = B2_VAL - B4_VAL
            Case Is < 13
                BINARY_BARRIER_OPTION_FUNC = A2_VAL - A4_VAL
            Case Is = 13
                BINARY_BARRIER_OPTION_FUNC = B3_VAL
            Case Is = 14
                BINARY_BARRIER_OPTION_FUNC = B3_VAL
            Case Is = 15
                BINARY_BARRIER_OPTION_FUNC = A3_VAL
            Case Is = 16
                BINARY_BARRIER_OPTION_FUNC = A1_VAL
            Case Is = 17
                BINARY_BARRIER_OPTION_FUNC = B2_VAL - B3_VAL + B4_VAL
            Case Is = 18
                BINARY_BARRIER_OPTION_FUNC = B1_VAL - B2_VAL + B4_VAL
            Case Is = 19
                BINARY_BARRIER_OPTION_FUNC = A2_VAL - A3_VAL + A4_VAL
            Case Is = 20
                BINARY_BARRIER_OPTION_FUNC = A1_VAL - A2_VAL + A3_VAL
            Case Is = 21
                BINARY_BARRIER_OPTION_FUNC = B1_VAL - B3_VAL
            Case Is = 22
                BINARY_BARRIER_OPTION_FUNC = 0
            Case Is = 23
                BINARY_BARRIER_OPTION_FUNC = A1_VAL - A3_VAL
            Case Is = 24
               BINARY_BARRIER_OPTION_FUNC = 0
            Case Is = 25
                BINARY_BARRIER_OPTION_FUNC = B1_VAL - B2_VAL + B3_VAL - B4_VAL
            Case Is = 26
                BINARY_BARRIER_OPTION_FUNC = B2_VAL - B4_VAL
            Case Is = 27
                BINARY_BARRIER_OPTION_FUNC = A1_VAL - A2_VAL + A3_VAL - A4_VAL
            Case Is = 28
                BINARY_BARRIER_OPTION_FUNC = A2_VAL - A4_VAL
        End Select
    
    ElseIf STRIKE < BARRIER Then
        Select Case OPTION_FLAG
            Case Is < 5
                BINARY_BARRIER_OPTION_FUNC = A5_VAL
            Case Is < 7
                BINARY_BARRIER_OPTION_FUNC = B2_VAL + B4_VAL
            Case Is < 9
                BINARY_BARRIER_OPTION_FUNC = A2_VAL + A4_VAL
            Case Is < 11
                BINARY_BARRIER_OPTION_FUNC = B2_VAL - B4_VAL
            Case Is < 13
                BINARY_BARRIER_OPTION_FUNC = A2_VAL - A4_VAL
            Case Is = 13
                BINARY_BARRIER_OPTION_FUNC = B1_VAL - B2_VAL + B4_VAL
            Case Is = 14
                BINARY_BARRIER_OPTION_FUNC = B2_VAL - B3_VAL + B4_VAL
            Case Is = 15
                BINARY_BARRIER_OPTION_FUNC = A1_VAL - A2_VAL + A4_VAL
            Case Is = 16
                BINARY_BARRIER_OPTION_FUNC = A2_VAL - A3_VAL + A4_VAL
            Case Is = 17
                BINARY_BARRIER_OPTION_FUNC = B1_VAL
            Case Is = 18
                BINARY_BARRIER_OPTION_FUNC = B3_VAL
            Case Is = 19
                BINARY_BARRIER_OPTION_FUNC = A1_VAL
            Case Is = 20
                BINARY_BARRIER_OPTION_FUNC = A3_VAL
            Case Is = 21
                BINARY_BARRIER_OPTION_FUNC = B2_VAL - B4_VAL
            Case Is = 22
                BINARY_BARRIER_OPTION_FUNC = B1_VAL - B2_VAL + B3_VAL - B4_VAL
            Case Is = 23
                BINARY_BARRIER_OPTION_FUNC = A2_VAL - A4_VAL
            Case Is = 24
                BINARY_BARRIER_OPTION_FUNC = A1_VAL - A2_VAL + A3_VAL - A4_VAL
            Case Is = 25
                BINARY_BARRIER_OPTION_FUNC = 0
            Case Is = 26
                BINARY_BARRIER_OPTION_FUNC = B1_VAL - B3_VAL
            Case Is = 27
                BINARY_BARRIER_OPTION_FUNC = 0
            Case Is = 28
                BINARY_BARRIER_OPTION_FUNC = A1_VAL - A3_VAL
        End Select
    End If
    
Exit Function
ERROR_LABEL:
BINARY_BARRIER_OPTION_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : KNOCK_OUT_CALL_OPTION_SIMULATION_FUNC
'DESCRIPTION   : KNOCK_OUT_CALL_OPTION_SIMULATION_FUNC
'LIBRARY       : DERIVATIVES
'GROUP         : BARRIER
'ID            : 009
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009

'************************************************************************************
'************************************************************************************

Function KNOCK_OUT_CALL_OPTION_SIMULATION_FUNC(ByVal nLOOPS As Long, _
ByVal DELTA As Double, _
ByVal SPOT As Double, _
ByVal STRIKE As Double, _
ByVal EXPIRATION As Double, _
ByVal RATE As Double, _
ByVal BARRIER As Double, _
ByVal SIGMA As Double, _
Optional ByVal OUTPUT As Integer = 0)

' nLOOPS = number of Monte Carlo replications
' DELTA = partition of time

Dim i As Long
Dim j As Long

Dim TEMP_RND As Double
Dim TEMP_SUM As Double
Dim TEMP_PRICE As Double

Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

ReDim TEMP_MATRIX(0 To nLOOPS, 1 To 3)

TEMP_MATRIX(0, 1) = "LOOP"
TEMP_MATRIX(0, 2) = "PRICE"
TEMP_MATRIX(0, 3) = "CUMULATIVE"

TEMP_SUM = 0
For i = 1 To nLOOPS
    TEMP_PRICE = SPOT
    j = 0
    Do While (TEMP_PRICE < BARRIER) And (j < DELTA)
        j = j + 1
        TEMP_RND = RANDOM_NORMAL_FUNC(0, 1, 0)
        TEMP_PRICE = TEMP_PRICE * Exp(((RATE - 0.5 * SIGMA * SIGMA) * _
        EXPIRATION / DELTA) + SIGMA * Sqr(EXPIRATION / DELTA) * TEMP_RND)
    Loop
    
    If (j = DELTA) And (TEMP_PRICE < BARRIER) Then: _
    TEMP_SUM = TEMP_SUM + MAXIMUM_FUNC(0, TEMP_PRICE - i)
    
    TEMP_MATRIX(i, 1) = i
    TEMP_MATRIX(i, 2) = TEMP_PRICE
    TEMP_MATRIX(i, 3) = TEMP_SUM
Next i

TEMP_SUM = TEMP_SUM / nLOOPS

Select Case OUTPUT
Case 0
    KNOCK_OUT_CALL_OPTION_SIMULATION_FUNC = Exp(-RATE * EXPIRATION) * TEMP_SUM
Case Else
    KNOCK_OUT_CALL_OPTION_SIMULATION_FUNC = TEMP_MATRIX
End Select

Exit Function
ERROR_LABEL:
KNOCK_OUT_CALL_OPTION_SIMULATION_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : DISCRETE_BARRIER_MONITORING_ADJ_FUNC
'DESCRIPTION   : Discrete barrier monitoring adjustment
'LIBRARY       : DERIVATIVES
'GROUP         : BARRIER
'ID            : 010
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009

'************************************************************************************
'************************************************************************************

Function DISCRETE_BARRIER_MONITORING_ADJ_FUNC(ByVal SPOT As Double, _
ByVal BARRIER As Double, _
ByVal SIGMA As Double, _
ByVal DELTA As Double, _
Optional ByVal ADJ_FACTOR As Double = 0.5826)

On Error GoTo ERROR_LABEL

    If BARRIER > SPOT Then
        DISCRETE_BARRIER_MONITORING_ADJ_FUNC = BARRIER * Exp(ADJ_FACTOR * _
        SIGMA * Sqr(DELTA))
    ElseIf BARRIER < SPOT Then
        DISCRETE_BARRIER_MONITORING_ADJ_FUNC = BARRIER * Exp(-1 * ADJ_FACTOR * _
        SIGMA * Sqr(DELTA))
    End If

Exit Function
ERROR_LABEL:
DISCRETE_BARRIER_MONITORING_ADJ_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : BARRIER_MONITORING_COUNT_BASIS_FUNC
'DESCRIPTION   : Barrier monitoring count basis
'LIBRARY       : DERIVATIVES
'GROUP         : BARRIER
'ID            : 011
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009

'************************************************************************************
'************************************************************************************

Function BARRIER_MONITORING_COUNT_BASIS_FUNC(Optional ByVal VERSION As Integer = 0)
On Error GoTo ERROR_LABEL
    Select Case VERSION
        Case 0 '"0", "C", "Continuously"
            BARRIER_MONITORING_COUNT_BASIS_FUNC = 0
        Case 1 '"1", "BARRIER", "Hourly"
            BARRIER_MONITORING_COUNT_BASIS_FUNC = 1 / (24 * 365)
        Case 2 '"2", "D", "Daily"
            BARRIER_MONITORING_COUNT_BASIS_FUNC = 1 / 365
        Case 3 ' "3", "W", "Weekly"
            BARRIER_MONITORING_COUNT_BASIS_FUNC = 1 / 52
        Case Else '"4", "M", "Monthly"
            BARRIER_MONITORING_COUNT_BASIS_FUNC = 1 / 12
    End Select
Exit Function
ERROR_LABEL:
BARRIER_MONITORING_COUNT_BASIS_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : BARRIER_BINARY_ETA_FACTOR_FUNC
'DESCRIPTION   : BINARY ETA FACTOR
'LIBRARY       : DERIVATIVES
'GROUP         : BARRIER
'ID            : 012
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009

'************************************************************************************
'************************************************************************************

Function BARRIER_BINARY_ETA_FACTOR_FUNC(ByVal VERSION As Integer)

On Error GoTo ERROR_LABEL

If VERSION = 1 Then: _
    BARRIER_BINARY_ETA_FACTOR_FUNC = 1 ' Down-and-in cash-(at-hit)-or-nothing (SPOT > BARRIER)
If VERSION = 2 Then: _
    BARRIER_BINARY_ETA_FACTOR_FUNC = -1 ' Up-and-in cash-(at-hit)-or-nothing (SPOT < BARRIER)
If VERSION = 3 Then: _
    BARRIER_BINARY_ETA_FACTOR_FUNC = 1 ' Down-and-in asset-(at-hit)-or-nothing (CASH = BARRIER)
    '(SPOT > BARRIER)
If VERSION = 4 Then: _
    BARRIER_BINARY_ETA_FACTOR_FUNC = -1 ' Up-and-in asset-(at-hit)-or-nothing (CASH = BARRIER)
    '(SPOT < BARRIER)
If VERSION = 5 Then: _
    BARRIER_BINARY_ETA_FACTOR_FUNC = 1 ' Down-and-in cash-(at-expiry)-or-nothing (SPOT > BARRIER)
If VERSION = 6 Then: _
    BARRIER_BINARY_ETA_FACTOR_FUNC = -1 ' Up-and-in cash-(at-expiry)-or-nothing (SPOT < BARRIER)
If VERSION = 7 Then: _
    BARRIER_BINARY_ETA_FACTOR_FUNC = 1 ' Down-and-in asset-(at-expiry)-or-nothing (SPOT > BARRIER)
If VERSION = 8 Then: _
    BARRIER_BINARY_ETA_FACTOR_FUNC = -1 ' Up-and-in asset-(at-expiry)-or-nothing (SPOT < BARRIER)
If VERSION = 9 Then: _
    BARRIER_BINARY_ETA_FACTOR_FUNC = 1 ' Down-and-out cash-(at-expiry)-or-nothing (SPOT > BARRIER)
If VERSION = 10 Then: _
    BARRIER_BINARY_ETA_FACTOR_FUNC = -1 ' Up-and-out cash-(at-expiry)-or-nothing (SPOT < BARRIER)
If VERSION = 11 Then: _
    BARRIER_BINARY_ETA_FACTOR_FUNC = 1 ' Down-and-out asset-(at-expiry)-or-nothing (SPOT > BARRIER)
If VERSION = 12 Then: _
    BARRIER_BINARY_ETA_FACTOR_FUNC = -1 ' Up-and-out asset-(at-expiry)-or-nothing (SPOT < BARRIER)
If VERSION = 13 Then: _
    BARRIER_BINARY_ETA_FACTOR_FUNC = 1 ' Down-and-in cash-(at-expiry)-or-nothing call (SPOT > BARRIER)
If VERSION = 14 Then: _
    BARRIER_BINARY_ETA_FACTOR_FUNC = -1 ' Up-and-in cash-(at-expiry)-or-nothing call (SPOT < BARRIER)
If VERSION = 15 Then: _
    BARRIER_BINARY_ETA_FACTOR_FUNC = 1 ' Down-and-in asset-(at-expiry)-or-nothing call (SPOT > BARRIER)
If VERSION = 16 Then: _
    BARRIER_BINARY_ETA_FACTOR_FUNC = -1 ' Up-and-in asset-(at-expiry)-or-nothing call (SPOT < BARRIER)
If VERSION = 17 Then: _
    BARRIER_BINARY_ETA_FACTOR_FUNC = 1 ' Down-and-in cash-(at-expiry)-or-nothing put (SPOT > BARRIER)
If VERSION = 18 Then: _
    BARRIER_BINARY_ETA_FACTOR_FUNC = -1 ' Up-and-in cash-(at-expiry)-or-nothing put (SPOT < BARRIER)
If VERSION = 19 Then: _
    BARRIER_BINARY_ETA_FACTOR_FUNC = 1 ' Down-and-in asset-(at-expiry)-or-nothing put (SPOT > BARRIER)
If VERSION = 20 Then: _
    BARRIER_BINARY_ETA_FACTOR_FUNC = -1 ' Up-and-in asset-(at-expiry)-or-nothing put (SPOT < BARRIER)
If VERSION = 21 Then: _
    BARRIER_BINARY_ETA_FACTOR_FUNC = 1 ' Down-and-out cash-(at-expiry)-or-nothing call (SPOT > BARRIER)
If VERSION = 22 Then: _
    BARRIER_BINARY_ETA_FACTOR_FUNC = -1 ' Up-and-out cash-(at-expiry)-or-nothing call (SPOT < BARRIER)
If VERSION = 23 Then: _
    BARRIER_BINARY_ETA_FACTOR_FUNC = 1 ' Down-and-out asset-(at-expiry)-or-nothing call (SPOT > BARRIER)
If VERSION = 24 Then: _
    BARRIER_BINARY_ETA_FACTOR_FUNC = -1 ' Up-and-out asset-(at-expiry)-or-nothing call (SPOT < BARRIER)
If VERSION = 25 Then: _
    BARRIER_BINARY_ETA_FACTOR_FUNC = 1 ' Down-and-out cash-(at-expiry)-or-nothing put (SPOT > BARRIER)
If VERSION = 26 Then: _
    BARRIER_BINARY_ETA_FACTOR_FUNC = -1 ' Up-and-out cash-(at-expiry)-or-nothing put (SPOT < BARRIER)
If VERSION = 27 Then: _
    BARRIER_BINARY_ETA_FACTOR_FUNC = 1 ' Down-and-out asset-(at-expiry)-or-nothing put (SPOT > BARRIER)
If VERSION = 28 Then: _
    BARRIER_BINARY_ETA_FACTOR_FUNC = -1 ' Up-and-out asset-(at-expiry)-or-nothing put (SPOT < BARRIER)

Exit Function
ERROR_LABEL:
BARRIER_BINARY_ETA_FACTOR_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : BARRIER_BINARY_PHI_FACTOR_FUNC
'DESCRIPTION   : BINARY PHI FACTOR
'LIBRARY       : DERIVATIVES
'GROUP         : BARRIER
'ID            : 013
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009

'************************************************************************************
'************************************************************************************


Function BARRIER_BINARY_PHI_FACTOR_FUNC(ByVal VERSION As Integer)

On Error GoTo ERROR_LABEL

If VERSION = 1 Then: _
    BARRIER_BINARY_PHI_FACTOR_FUNC = 0 ' Down-and-in cash-(at-hit)-or-nothing (SPOT > BARRIER)
If VERSION = 2 Then: _
    BARRIER_BINARY_PHI_FACTOR_FUNC = 0 ' Up-and-in cash-(at-hit)-or-nothing (SPOT < BARRIER)
If VERSION = 3 Then: _
    BARRIER_BINARY_PHI_FACTOR_FUNC = 0 ' Down-and-in asset-(at-hit)-or-nothing (CASH = BARRIER)
    '(SPOT > BARRIER)
If VERSION = 4 Then: _
    BARRIER_BINARY_PHI_FACTOR_FUNC = 0 ' Up-and-in asset-(at-hit)-or-nothing (CASH = BARRIER)
    '(SPOT < BARRIER)
If VERSION = 5 Then: _
    BARRIER_BINARY_PHI_FACTOR_FUNC = -1 ' Down-and-in cash-(at-expiry)-or-nothing (SPOT > BARRIER)
If VERSION = 6 Then: _
    BARRIER_BINARY_PHI_FACTOR_FUNC = 1 ' Up-and-in cash-(at-expiry)-or-nothing (SPOT < BARRIER)
If VERSION = 7 Then: _
    BARRIER_BINARY_PHI_FACTOR_FUNC = -1 ' Down-and-in asset-(at-expiry)-or-nothing (SPOT > BARRIER)
If VERSION = 8 Then: _
    BARRIER_BINARY_PHI_FACTOR_FUNC = 1 ' Up-and-in asset-(at-expiry)-or-nothing (SPOT < BARRIER)
If VERSION = 9 Then: _
    BARRIER_BINARY_PHI_FACTOR_FUNC = 1 ' Down-and-out cash-(at-expiry)-or-nothing (SPOT > BARRIER)
If VERSION = 10 Then: _
    BARRIER_BINARY_PHI_FACTOR_FUNC = -1 ' Up-and-out cash-(at-expiry)-or-nothing (SPOT < BARRIER)
If VERSION = 11 Then: _
    BARRIER_BINARY_PHI_FACTOR_FUNC = 1 ' Down-and-out asset-(at-expiry)-or-nothing (SPOT > BARRIER)
If VERSION = 12 Then: _
    BARRIER_BINARY_PHI_FACTOR_FUNC = -1 ' Up-and-out asset-(at-expiry)-or-nothing (SPOT < BARRIER)
If VERSION = 13 Then: _
    BARRIER_BINARY_PHI_FACTOR_FUNC = 1 ' Down-and-in cash-(at-expiry)-or-nothing call (SPOT > BARRIER)
If VERSION = 14 Then: _
    BARRIER_BINARY_PHI_FACTOR_FUNC = 1 ' Up-and-in cash-(at-expiry)-or-nothing call (SPOT < BARRIER)
If VERSION = 15 Then: _
    BARRIER_BINARY_PHI_FACTOR_FUNC = 1 ' Down-and-in asset-(at-expiry)-or-nothing call (SPOT > BARRIER)
If VERSION = 16 Then: _
    BARRIER_BINARY_PHI_FACTOR_FUNC = 1 ' Up-and-in asset-(at-expiry)-or-nothing call (SPOT < BARRIER)
If VERSION = 17 Then: _
    BARRIER_BINARY_PHI_FACTOR_FUNC = -1 ' Down-and-in cash-(at-expiry)-or-nothing put (SPOT > BARRIER)
If VERSION = 18 Then: _
    BARRIER_BINARY_PHI_FACTOR_FUNC = -1 ' Up-and-in cash-(at-expiry)-or-nothing put (SPOT < BARRIER)
If VERSION = 19 Then: _
    BARRIER_BINARY_PHI_FACTOR_FUNC = -1 ' Down-and-in asset-(at-expiry)-or-nothing put (SPOT > BARRIER)
If VERSION = 20 Then: _
    BARRIER_BINARY_PHI_FACTOR_FUNC = -1 ' Up-and-in asset-(at-expiry)-or-nothing put (SPOT < BARRIER)
If VERSION = 21 Then: _
    BARRIER_BINARY_PHI_FACTOR_FUNC = 1 ' Down-and-out cash-(at-expiry)-or-nothing call (SPOT > BARRIER)
If VERSION = 22 Then: _
    BARRIER_BINARY_PHI_FACTOR_FUNC = 1 ' Up-and-out cash-(at-expiry)-or-nothing call (SPOT < BARRIER)
If VERSION = 23 Then: _
    BARRIER_BINARY_PHI_FACTOR_FUNC = 1 ' Down-and-out asset-(at-expiry)-or-nothing call (SPOT > BARRIER)
If VERSION = 24 Then: _
    BARRIER_BINARY_PHI_FACTOR_FUNC = 1 ' Up-and-out asset-(at-expiry)-or-nothing call (SPOT < BARRIER)
If VERSION = 25 Then: _
    BARRIER_BINARY_PHI_FACTOR_FUNC = -1 ' Down-and-out cash-(at-expiry)-or-nothing put (SPOT > BARRIER)
If VERSION = 26 Then: _
    BARRIER_BINARY_PHI_FACTOR_FUNC = -1 ' Up-and-out cash-(at-expiry)-or-nothing put (SPOT < BARRIER)
If VERSION = 27 Then: _
    BARRIER_BINARY_PHI_FACTOR_FUNC = -1 ' Down-and-out asset-(at-expiry)-or-nothing put (SPOT > BARRIER)
If VERSION = 28 Then: _
    BARRIER_BINARY_PHI_FACTOR_FUNC = -1 ' Up-and-out asset-(at-expiry)-or-nothing put (SPOT < BARRIER)

Exit Function
ERROR_LABEL:
BARRIER_BINARY_PHI_FACTOR_FUNC = Err.number
End Function

'---------------------------------------------------------------------------------
'----------------------------About Barrier Options--------------------------------
'---------------------------------------------------------------------------------

'Description

'Barrier options are similar to standard options except that they are
'extinguished or activated when the underlying asset price reaches a
'predetermined barrier or boundary price.

'The payoff of a standard European style option is based on the price of
'the underlying asset on the expiration date. A standard option is
'path-independent since it does not matter what path the underlying asset
'took during the option life.

'Barrier options are path-dependent since they are dependent on the price
'movement of the underlying asset. A knock-out option will expire early if
'the barrier price is reached whereas a knock-in option will come into
'existence if the barrier price is reached. As with average options, a
'monitoring frequency is defined as part of the option which specifies how
'often the price is checked for breach of the barrier. The frequency is
'normally continuous but could be hourly, daily, etc.

'Types of barrier options:
'1) Down and Out
    'The option is canceled or knocked-out if the asset falls to a '
    'predetermined boundary price.
'2) Down and In
    'The option is activated or knocked-in if the asset falls to a
    'predetermined boundary price.
'3) Up and Out
    'The option is canceled or knocked-out if the asset rises to a
    'predetermined boundary price.
'4) Up and In
    'The option is activated or knocked-in if the asset rises to a
    'predetermined boundary price.

'----------------------------------------------------------------------------------
'Benefits
'----------------------------------------------------------------------------------
'The premium for barrier options is lower than standard options, as the
'barrier option will have value within a smaller price range than the
'standard option.

'The owner of a barrier option loses some of the traditional option value and
'therefore it should sell at a lower price than a standard option.

'The seller of the barrier lowers his exposure or risk, relative to a standard
'option. Some barrier options offer a rebate; should the option be knocked-out,
'the holder would receive a predefined payoff. This feature is less common.
'Obviously a barrier option with rebate has more value than one without.

'Features
'What do you have when you buy a down-and-out call and a down-and-in call,
'assuming no rebate? A standard call!

'Barrier options are hybrids: they are European in that they could have a
'payoff at expiration but they are American in that they may exercised
'(extinguished) prior to expiration.

'The lower cost of the barrier option makes it one of the most popular of the exotic
'options for hedging purposes.

'Speculators are able to gain greater leverage with barrier options for the same
'dollar amount.
'Intrinsic Value Formula

'The payoff at expiration is:
'Call: max {0, U E}
'Put: max {0, E U}

'Where U is the underlying price and E is the exercise or strike price, additional
'condition is that D&O and U&O are worth R or rebate if barrier is reached.

'Aliases
'Barrier options are also known as down-and-outs, down-and-ins, up-and-outs,
'up-and-ins, knock-outs, knock-ins, kick-outs, kick-ins, ins, outs, exploding
'options, extinguishing options and trigger options.

'Uses
'1) A bank may wish to purchase an at-the-money 9-month Nikkei call option struck
'at 17,000 with a down-and-out barrier price of 16,000. If the price of the
'Nikkei falls to 16,000 or below, during the 9-month period, the bank will no longer
'have the benefit of Nikkei price appreciation since the call option will have been
'knocked out.

'2) An airline is concerned that events in the Middle East might drive up the price
'of fuel. An up-and-in call would allow the airline to buy crude oil futures at a
'fixed price if some knock-in boundary price is reached. The price of the U&i call
'would be less than a standard call with the same expiration and exercise price so
'it might be viewed as a cost effective hedging instrument.

'---------------------------------------------------------------------------------
'---------------------------------------------------------------------------------
'---------------------------------------------------------------------------------
