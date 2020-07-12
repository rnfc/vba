Attribute VB_Name = "FINAN_DERIV_EXCHANGE_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : EXCHANGE_ONE_ASSET_OPTION_FUNC
'DESCRIPTION   : Option to exchange one asset for another
'LIBRARY       : DERIVATIVES
'GROUP         : EXCHANGE
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function EXCHANGE_ONE_ASSET_OPTION_FUNC(ByVal SPOT_A As Double, _
ByVal SPOT_B As Double, _
ByVal QUANTITY_A As Double, _
ByVal QUANTITY_B As Double, _
ByVal EXPIRATION As Double, _
ByVal RATE As Double, _
ByVal CARRY_COST_A As Double, _
ByVal CARRY_COST_B As Double, _
ByVal SIGMA_A As Double, _
ByVal SIGMA_B As Double, _
ByVal RHO_VAL As Double, _
Optional ByVal EXERCISE_TYPE As Integer = 0, _
Optional ByVal CND_TYPE As Integer = 0)

'RHO = Correlation(A,B)

Dim D1_VAL As Double
Dim D2_VAL As Double
Dim TEMP_SIGMA As Double
 
On Error GoTo ERROR_LABEL

Select Case EXERCISE_TYPE
    Case 0 ', "EUROPEAN", "EURO", "E"

       TEMP_SIGMA = Sqr(SIGMA_A ^ 2 + SIGMA_B ^ 2 - 2 * RHO_VAL * SIGMA_A * SIGMA_B)
       
       D1_VAL = (Log(QUANTITY_A * SPOT_A / (QUANTITY_B * SPOT_B)) + _
       (CARRY_COST_A - CARRY_COST_B + TEMP_SIGMA ^ 2 / 2) * EXPIRATION) / _
       (TEMP_SIGMA * Sqr(EXPIRATION))
       
       D2_VAL = D1_VAL - TEMP_SIGMA * Sqr(EXPIRATION)
    
       EXCHANGE_ONE_ASSET_OPTION_FUNC = QUANTITY_A * SPOT_A * Exp((CARRY_COST_A - _
       RATE) * EXPIRATION) * CND_FUNC(D1_VAL, CND_TYPE) - QUANTITY_B * SPOT_B * _
       Exp((CARRY_COST_B - RATE) * EXPIRATION) * CND_FUNC(D2_VAL, CND_TYPE)
    
    Case Else '1 ', "AMERICAN", "AMER", "A"
    
        TEMP_SIGMA = Sqr(SIGMA_A ^ 2 + SIGMA_B ^ 2 - 2 * RHO_VAL * SIGMA_A * SIGMA_B)
        
        EXCHANGE_ONE_ASSET_OPTION_FUNC = APPROXIMATION_AMERICAN_OPTION_FUNC(QUANTITY_A * SPOT_A, _
        QUANTITY_B * SPOT_B, EXPIRATION, RATE - CARRY_COST_B, CARRY_COST_A - _
        CARRY_COST_B, TEMP_SIGMA, 1, 0)
End Select
    
Exit Function
ERROR_LABEL:
EXCHANGE_ONE_ASSET_OPTION_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : EXCHANGE_OPTION_EXCHANGE_OPTION_FUNC
'DESCRIPTION   : Exchange options on exchange options
'LIBRARY       : DERIVATIVES
'GROUP         : EXCHANGE
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function EXCHANGE_OPTION_EXCHANGE_OPTION_FUNC(ByVal SPOT_A As Double, _
ByVal SPOT_B As Double, _
ByVal QUANTITY_B As Double, _
ByVal MATURITY As Double, _
ByVal EXPIRATION As Double, _
ByVal RATE As Double, _
ByVal CARRY_COST_A As Double, _
ByVal CARRY_COST_B As Double, _
ByVal SIGMA_A As Double, _
ByVal SIGMA_B As Double, _
ByVal RHO_VAL As Double, _
Optional ByVal OPTION_FLAG As Integer = 1, _
Optional ByVal CND_TYPE As Integer = 0, _
Optional ByVal CBND_TYPE As Integer = 0)
    

'MATURITY: Time to maturity
'EXPIRATION: Time to maturity underlying option

 
'[OPTION_FLAG: 1] Option to exchange Q * S2 for the option to
'exchange S2 for S1

'[OPTION_FLAG: 2] Option to exchange the option to exchange S2 for S1
'in return for Q*S2

'[OPTION_FLAG: 3] Option to exchange Q*S2 for the option to exchange S1 for S2

'[OPTION_FLAG: 4] Option to exchange the option to exchange S1 for S2 in return
'for Q*S2

Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double

Dim DI_VAL As Double
Dim D1_VAL As Double
Dim D2_VAL As Double
Dim D3_VAL As Double
Dim D4_VAL As Double

Dim YI_VAL As Double
Dim Y1_VAL As Double
Dim Y2_VAL As Double
Dim Y3_VAL As Double
Dim Y4_VAL As Double

Dim ZI_VAL As Double
Dim Z1_VAL As Double
Dim Z2_VAL As Double

Dim TEMP_SIGMA As Double

Dim TEMP_FLAG As Integer
Dim tolerance As Double

On Error GoTo ERROR_LABEL

TEMP_SIGMA = Sqr(SIGMA_A ^ 2 + SIGMA_B ^ 2 - 2 * RHO_VAL * SIGMA_A * SIGMA_B)
ATEMP_VAL = SPOT_A * Exp((CARRY_COST_A - RATE) * (EXPIRATION - MATURITY)) / _
(SPOT_B * Exp((CARRY_COST_B - RATE) * (EXPIRATION - MATURITY)))

If OPTION_FLAG = 1 Or OPTION_FLAG = 2 Then
    TEMP_FLAG = 1
Else: TEMP_FLAG = 2
End If


ZI_VAL = ATEMP_VAL
If TEMP_FLAG = 1 Then
    Z1_VAL = (Log(ZI_VAL) + TEMP_SIGMA ^ 2 / 2 * (EXPIRATION - MATURITY)) / _
    (TEMP_SIGMA * Sqr(EXPIRATION - MATURITY))
    Z2_VAL = (Log(ZI_VAL) - TEMP_SIGMA ^ 2 / 2 * (EXPIRATION - MATURITY)) / _
    (TEMP_SIGMA * Sqr(EXPIRATION - MATURITY))
    YI_VAL = ZI_VAL * CND_FUNC(Z1_VAL, CND_TYPE) - CND_FUNC(Z2_VAL, CND_TYPE)
ElseIf TEMP_FLAG = 2 Then
    Z1_VAL = (-Log(ZI_VAL) + TEMP_SIGMA ^ 2 / 2 * (EXPIRATION - MATURITY)) / _
    (TEMP_SIGMA * Sqr(EXPIRATION - MATURITY))
    Z2_VAL = (-Log(ZI_VAL) - TEMP_SIGMA ^ 2 / 2 * (EXPIRATION - MATURITY)) / _
    (TEMP_SIGMA * Sqr(EXPIRATION - MATURITY))
    YI_VAL = CND_FUNC(Z1_VAL, CND_TYPE) - ZI_VAL * CND_FUNC(Z2_VAL, CND_TYPE)
End If

ZI_VAL = ATEMP_VAL
If TEMP_FLAG = 1 Then
    Z1_VAL = (Log(ZI_VAL) + TEMP_SIGMA ^ 2 / 2 * (EXPIRATION - MATURITY)) / _
    (TEMP_SIGMA * Sqr(EXPIRATION - MATURITY))
    DI_VAL = CND_FUNC(Z1_VAL, CND_TYPE)
ElseIf TEMP_FLAG = 2 Then
    Z2_VAL = (-Log(ZI_VAL) - TEMP_SIGMA ^ 2 / 2 * (EXPIRATION - MATURITY)) / _
    (TEMP_SIGMA * Sqr(EXPIRATION - MATURITY))
    DI_VAL = -CND_FUNC(Z2_VAL, CND_TYPE)
End If

tolerance = 0.00001

'---------------------------------------------------------------------------
'---------------------Numerical search algorithm to find critical price
'---------------------------------------------------------------------------

Do While Abs(YI_VAL - QUANTITY_B) > tolerance
    ZI_VAL = ZI_VAL - (YI_VAL - QUANTITY_B) / DI_VAL
    BTEMP_VAL = ZI_VAL
    If TEMP_FLAG = 1 Then
        Z1_VAL = (Log(ZI_VAL) + TEMP_SIGMA ^ 2 / 2 * (EXPIRATION - MATURITY)) / _
        (TEMP_SIGMA * Sqr(EXPIRATION - MATURITY))
        Z2_VAL = (Log(ZI_VAL) - TEMP_SIGMA ^ 2 / 2 * (EXPIRATION - MATURITY)) / _
        (TEMP_SIGMA * Sqr(EXPIRATION - MATURITY))
        YI_VAL = ZI_VAL * CND_FUNC(Z1_VAL, CND_TYPE) - CND_FUNC(Z2_VAL, CND_TYPE)
    ElseIf TEMP_FLAG = 2 Then
        Z1_VAL = (-Log(ZI_VAL) + TEMP_SIGMA ^ 2 / 2 * (EXPIRATION - MATURITY)) / _
        (TEMP_SIGMA * Sqr(EXPIRATION - MATURITY))
        Z2_VAL = (-Log(ZI_VAL) - TEMP_SIGMA ^ 2 / 2 * (EXPIRATION - MATURITY)) / _
        (TEMP_SIGMA * Sqr(EXPIRATION - MATURITY))
        YI_VAL = CND_FUNC(Z1_VAL, CND_TYPE) - ZI_VAL * CND_FUNC(Z2_VAL, CND_TYPE)
    End If
    
    ZI_VAL = BTEMP_VAL
    If TEMP_FLAG = 1 Then
        Z1_VAL = (Log(ZI_VAL) + TEMP_SIGMA ^ 2 / 2 * (EXPIRATION - MATURITY)) / _
        (TEMP_SIGMA * Sqr(EXPIRATION - MATURITY))
        DI_VAL = CND_FUNC(Z1_VAL, CND_TYPE)
    ElseIf TEMP_FLAG = 2 Then
        Z2_VAL = (-Log(ZI_VAL) - TEMP_SIGMA ^ 2 / 2 * (EXPIRATION - MATURITY)) / _
        (TEMP_SIGMA * Sqr(EXPIRATION - MATURITY))
        DI_VAL = -CND_FUNC(Z2_VAL, CND_TYPE)
    End If
Loop

BTEMP_VAL = ZI_VAL

D1_VAL = (Log(SPOT_A / (BTEMP_VAL * SPOT_B)) + (CARRY_COST_A - CARRY_COST_B + _
TEMP_SIGMA ^ 2 / 2) * MATURITY) / (TEMP_SIGMA * Sqr(MATURITY))

D2_VAL = D1_VAL - TEMP_SIGMA * Sqr(MATURITY)

D3_VAL = (Log((BTEMP_VAL * SPOT_B) / SPOT_A) + (CARRY_COST_B - CARRY_COST_A + _
TEMP_SIGMA ^ 2 / 2) * MATURITY) / (TEMP_SIGMA * Sqr(MATURITY))

D4_VAL = D3_VAL - TEMP_SIGMA * Sqr(MATURITY)

Y1_VAL = (Log(SPOT_A / SPOT_B) + (CARRY_COST_A - CARRY_COST_B + _
TEMP_SIGMA ^ 2 / 2) * EXPIRATION) / (TEMP_SIGMA * Sqr(EXPIRATION))

Y2_VAL = Y1_VAL - TEMP_SIGMA * Sqr(EXPIRATION)

Y3_VAL = (Log(SPOT_B / SPOT_A) + (CARRY_COST_B - CARRY_COST_A + _
TEMP_SIGMA ^ 2 / 2) * EXPIRATION) / (TEMP_SIGMA * Sqr(EXPIRATION))

Y4_VAL = Y3_VAL - TEMP_SIGMA * Sqr(EXPIRATION)

    Select Case OPTION_FLAG
    Case 1
        EXCHANGE_OPTION_EXCHANGE_OPTION_FUNC = -SPOT_B * Exp((CARRY_COST_B - RATE) * _
        EXPIRATION) * CBND_FUNC(D2_VAL, Y2_VAL, Sqr(MATURITY / EXPIRATION), _
        CND_TYPE, CBND_TYPE) + SPOT_A * _
        Exp((CARRY_COST_A - RATE) * EXPIRATION) * CBND_FUNC(D1_VAL, Y1_VAL, Sqr(MATURITY / _
        EXPIRATION), CND_TYPE, CBND_TYPE) - QUANTITY_B * SPOT_B * _
        Exp((CARRY_COST_B - RATE) * _
        MATURITY) * CND_FUNC(D2_VAL, CND_TYPE)
    Case 2
        EXCHANGE_OPTION_EXCHANGE_OPTION_FUNC = SPOT_B * Exp((CARRY_COST_B - RATE) * _
        EXPIRATION) * CBND_FUNC(D3_VAL, Y2_VAL, -Sqr(MATURITY / EXPIRATION), _
        CND_TYPE, CBND_TYPE) - SPOT_A * _
        Exp((CARRY_COST_A - RATE) * EXPIRATION) * _
        CBND_FUNC(D4_VAL, Y1_VAL, -Sqr(MATURITY / _
        EXPIRATION), CND_TYPE, CBND_TYPE) + QUANTITY_B * SPOT_B * _
        Exp((CARRY_COST_B - RATE) * _
        MATURITY) * CND_FUNC(D3_VAL, CND_TYPE)
    Case 3
        EXCHANGE_OPTION_EXCHANGE_OPTION_FUNC = SPOT_B * Exp((CARRY_COST_B - RATE) * _
        EXPIRATION) * CBND_FUNC(D3_VAL, Y3_VAL, Sqr(MATURITY / EXPIRATION), _
        CND_TYPE, CBND_TYPE) - SPOT_A * _
        Exp((CARRY_COST_A - RATE) * EXPIRATION) * _
        CBND_FUNC(D4_VAL, Y4_VAL, Sqr(MATURITY / _
        EXPIRATION), CND_TYPE, CBND_TYPE) - QUANTITY_B * SPOT_B * _
        Exp((CARRY_COST_B - RATE) * _
        MATURITY) * CND_FUNC(D3_VAL, CND_TYPE)
    Case Else '4
        EXCHANGE_OPTION_EXCHANGE_OPTION_FUNC = -SPOT_B * Exp((CARRY_COST_B - RATE) * _
        EXPIRATION) * CBND_FUNC(D2_VAL, Y3_VAL, -Sqr(MATURITY / EXPIRATION), _
        CND_TYPE, CBND_TYPE) + SPOT_A * _
        Exp((CARRY_COST_A - RATE) * EXPIRATION) * CBND_FUNC(D1_VAL, Y4_VAL, -Sqr(MATURITY / _
        EXPIRATION), CND_TYPE, CBND_TYPE) + QUANTITY_B * SPOT_B * _
        Exp((CARRY_COST_B - RATE) * MATURITY) _
        * CND_FUNC(D2_VAL, CND_TYPE)
    End Select
    
Exit Function
ERROR_LABEL:
EXCHANGE_OPTION_EXCHANGE_OPTION_FUNC = Err.number
End Function
