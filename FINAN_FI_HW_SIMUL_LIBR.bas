Attribute VB_Name = "FINAN_FI_HW_SIMUL_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : HW_SWAPTION_MC_FUNC
'DESCRIPTION   : Swaption pricing in Hull White model
'dr_t = (mu_t - a*r_t)dt + sigma*dW_t
'LIBRARY       : FIXED_INCOME
'GROUP         : HW_SIMUL
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'**********************************************************************************
'**********************************************************************************

Function HW_SWAPTION_MC_FUNC(ByVal HW_RATE As Double, _
ByVal HW_VOLATILITY As Double, _
ByVal FLAT_FORWARD_RATE As Double, _
ByVal SWAP_START_TENOR As Double, _
ByVal SWAP_END_TENOR As Double, _
ByVal DELTA As Double, _
ByVal STRIKE As Double, _
ByVal NOMINAL As Double, _
Optional ByVal nLOOPS As Single = 1000)

'HW_RATE --> Hull White Model Specification
'HW_VOLATILITY --> Hull White Model Specification

'FLAT_FORWARD_RATE --> It is assumed a flat forward term structure.
'Enter continously compounded forward rate

'Swaption Specification
'SWAP_START_TENOR, SWAP_END_TENOR
'DELTA, STRIKE, NOMINAL

Dim i As Long
Dim j As Long
Dim k As Long

Dim B_VAL As Double 'bondPrice0
Dim D_VAL As Double 'std_rt
Dim E_VAL As Double 'Ert
Dim R_VAL As Double
Dim S_VAL As Double
Dim T_VAL As Double
Dim V_VAL As Double 'Var_rt

Dim P0_VAL As Double
Dim PT_VAL As Double

Dim SR1_VAL As Double
Dim SR2_VAL As Double

Dim TEMP_VAL As Double

Dim ATEMP_SUM As Double
Dim BTEMP_SUM As Double

Dim PV_FIXED_VAL As Double
Dim PAYOFF_CALL_VAL As Double

Dim NORMAL_RAND_ARR As Variant

On Error GoTo ERROR_LABEL

'HW_RATE(a) 0.1
'HW_VOLATILITY(sigma) 0.04
'FlatForward 0.05
'Swap start time 1
'Swap End Time   6
'Tenor 0.5
'STRIKE 2
'Nominal amount  1000
'Swaption Price  69.32061459
 
Randomize
k = (SWAP_END_TENOR - SWAP_START_TENOR) / DELTA

If STRIKE = 0 Then
    STRIKE = FAIR_SWAP_RATE_FUNC(HW_RATE, HW_VOLATILITY, FLAT_FORWARD_RATE, _
             SWAP_START_TENOR, SWAP_END_TENOR, DELTA)
End If

ReDim FIXED_ARR(1 To k) 'store the fixed leg payments
ATEMP_SUM = SWAP_START_TENOR
BTEMP_SUM = 0
For i = 1 To k
  ATEMP_SUM = ATEMP_SUM + DELTA
  FIXED_ARR(i) = DELTA * STRIKE
  P0_VAL = DISCOUNT_BOND_FUNC(HW_RATE, HW_VOLATILITY, FLAT_FORWARD_RATE, FLAT_FORWARD_RATE, 0, ATEMP_SUM)
  BTEMP_SUM = BTEMP_SUM + P0_VAL * DELTA
Next i
SR1_VAL = (DISCOUNT_BOND_FUNC(HW_RATE, HW_VOLATILITY, FLAT_FORWARD_RATE, FLAT_FORWARD_RATE, 0, SWAP_START_TENOR) - P0_VAL) / BTEMP_SUM
FIXED_ARR(k) = FIXED_ARR(k) + 1


P0_VAL = DISCOUNT_BOND_FUNC(HW_RATE, HW_VOLATILITY, FLAT_FORWARD_RATE, FLAT_FORWARD_RATE, 0, SWAP_START_TENOR)
E_VAL = -MFN_FUNC(HW_RATE, HW_VOLATILITY, 0, SWAP_START_TENOR, SWAP_START_TENOR) + _
       ALPHA_T_FUNC(HW_RATE, HW_VOLATILITY, FLAT_FORWARD_RATE, SWAP_START_TENOR)
T_VAL = SWAP_START_TENOR
S_VAL = 0
V_VAL = (HW_VOLATILITY * HW_VOLATILITY / (2 * HW_RATE)) * (1 - Exp(-2 * HW_RATE * (T_VAL - S_VAL)))
 
'we make use of affine property which states
'that bond price is an affine fucntion of short rate
B_VAL = DISCOUNT_BOND_FUNC(HW_RATE, HW_VOLATILITY, FLAT_FORWARD_RATE, _
        FLAT_FORWARD_RATE, SWAP_START_TENOR, SWAP_END_TENOR)

'BT_VAL = BFN_FUNC(HW_RATE, SWAP_START_TENOR, MATURITY_VAL)
'B_VAL * Exp(BT_VAL * FLAT_FORWARD_RATE)
ReDim BTEMP_ARR(1 To k): ReDim ATEMP_ARR(1 To k)
ATEMP_SUM = SWAP_START_TENOR
For j = 1 To k
  ATEMP_SUM = ATEMP_SUM + DELTA
  ATEMP_ARR(j) = BFN_FUNC(HW_RATE, SWAP_START_TENOR, ATEMP_SUM)
  B_VAL = DISCOUNT_BOND_FUNC(HW_RATE, HW_VOLATILITY, _
               FLAT_FORWARD_RATE, FLAT_FORWARD_RATE, SWAP_START_TENOR, ATEMP_SUM)
  BTEMP_ARR(j) = B_VAL * Exp(ATEMP_ARR(j) * FLAT_FORWARD_RATE)
Next j

D_VAL = V_VAL ^ 0.5
ATEMP_SUM = 0
NORMAL_RAND_ARR = VECTOR_RANDOM_BOX_MULLER_FUNC(nLOOPS) 'VECTOR_RANDOM_BOX_MULLER_FUNC(nLOOPS)
NORMAL_RAND_ARR(1) = 1

PV_FIXED_VAL = 0
BTEMP_SUM = 0
PAYOFF_CALL_VAL = 0
For i = 1 To nLOOPS
  R_VAL = E_VAL + D_VAL * NORMAL_RAND_ARR(i)
  ATEMP_SUM = 0
  For j = 1 To k
    PT_VAL = BTEMP_ARR(j) * Exp(-ATEMP_ARR(j) * R_VAL)
    ATEMP_SUM = ATEMP_SUM + PT_VAL * DELTA
  Next j
  SR2_VAL = (1 - PT_VAL) / ATEMP_SUM
  TEMP_VAL = P0_VAL * NOMINAL * MAXIMUM_FUNC(SR2_VAL - SR1_VAL, 0) * ATEMP_SUM
  BTEMP_SUM = BTEMP_SUM + SR2_VAL
  PAYOFF_CALL_VAL = PAYOFF_CALL_VAL + TEMP_VAL
Next i
'PAYOFF_CALL_VAL / nLOOPS --> averagepayoffcall
'BTEMP_SUM / nLOOPS --> averageSR
HW_SWAPTION_MC_FUNC = Array(PAYOFF_CALL_VAL / nLOOPS, BTEMP_SUM / nLOOPS)

Exit Function
ERROR_LABEL:
HW_SWAPTION_MC_FUNC = Err.number
End Function

'**********************************************************************************
'**********************************************************************************
'FUNCTION      : HW_SWAPTION_MC_FUNC

'DESCRIPTION   : Bond Option pricing in Hull White model
'dr_t = (mu_t - a*r_t)dt + \sigma dW_t

'LIBRARY       : FIXED_INCOME
'GROUP         : HW_SIMUL
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'**********************************************************************************
'**********************************************************************************

Function HW_BOND_OPTION_MC_FUNC(ByVal HW_RATE As Double, _
ByVal HW_VOLATILITY As Double, _
ByVal FLAT_FORWARD_RATE As Double, _
ByVal OPTION_MATURITY As Double, _
ByVal BOND_MATURITY As Double, _
ByVal STRIKE As Double, _
Optional ByVal nLOOPS As Single = 5000, _
Optional ByVal OUTPUT As Integer = 0, _
Optional ByVal CND_TYPE As Integer = 0)

Dim i As Long
Dim BT_VAL As Double
Dim E_VAL As Double
Dim D_VAL As Double 'std_rt
Dim P_VAL As Double 'bondPrice0
Dim R_VAL As Double
Dim T_VAL As Double
Dim V_VAL As Double 'Var_rt
Dim PT_VAL As Double
Dim P0T_VAL As Double

Dim CALL_VAL As Double
Dim PUT_VAL As Double

Dim NORMAL_RAND_ARR As Variant
Dim CALL_PAYOFF_SUM_VAL As Double
Dim PUT_PAYOFF_SUM_VAL As Double

'-----------------------------------------
'HW_RATE(a) 0.1
'HW_VOLATILITY(sigma) 0.04
'FlatForward 0.05
'option maturity 5
'bond maturity   10
'STRIKE 0.9
'Forward Bond price  0.748903242
'-----------------------------------------
'Call Option 0.034965437
'Put Option  0.129355482
'-----------------------------------------
'Call Option 0.034789847
'Put Option  0.129524468
'-----------------------------------------

On Error GoTo ERROR_LABEL

If OUTPUT <> 0 Or nLOOPS = 1 Then 'Analytical Price
    P0T_VAL = DISCOUNT_BOND_FUNC(HW_RATE, HW_VOLATILITY, FLAT_FORWARD_RATE, _
              FLAT_FORWARD_RATE, OPTION_MATURITY, BOND_MATURITY)
    
    CALL_VAL = BOND_OPTION_DISCOUNT_FUNC(HW_RATE, HW_VOLATILITY, FLAT_FORWARD_RATE, _
               STRIKE, 0, OPTION_MATURITY, BOND_MATURITY, 1, CND_TYPE)
    
    PUT_VAL = BOND_OPTION_DISCOUNT_FUNC(HW_RATE, HW_VOLATILITY, FLAT_FORWARD_RATE, _
              STRIKE, 0, OPTION_MATURITY, BOND_MATURITY, -1, CND_TYPE)
              
    HW_BOND_OPTION_MC_FUNC = Array(CALL_VAL, PUT_VAL, P0T_VAL)
    Exit Function
End If

Randomize
CALL_PAYOFF_SUM_VAL = 0
PUT_PAYOFF_SUM_VAL = 0

P0T_VAL = DISCOUNT_BOND_FUNC(HW_RATE, HW_VOLATILITY, FLAT_FORWARD_RATE, _
          FLAT_FORWARD_RATE, 0, OPTION_MATURITY)

E_VAL = -MFN_FUNC(HW_RATE, HW_VOLATILITY, 0, OPTION_MATURITY, OPTION_MATURITY) + _
         ALPHA_T_FUNC(HW_RATE, HW_VOLATILITY, FLAT_FORWARD_RATE, OPTION_MATURITY)

V_VAL = (HW_VOLATILITY * HW_VOLATILITY / (2 * HW_RATE)) * _
         (1 - Exp(-2 * HW_RATE * (OPTION_MATURITY - 0)))

'we make use of affine property which states
'that bond price is an affine fucntion of short rate
'hence we
P_VAL = DISCOUNT_BOND_FUNC(HW_RATE, HW_VOLATILITY, FLAT_FORWARD_RATE, _
        FLAT_FORWARD_RATE, OPTION_MATURITY, BOND_MATURITY)
'BT_VAL = (1 - Exp(-HW_RATE * (BOND_MATURITY - OPTION_MATURITY))) / HW_RATE
BT_VAL = BFN_FUNC(HW_RATE, OPTION_MATURITY, BOND_MATURITY)
T_VAL = P_VAL * Exp(BT_VAL * FLAT_FORWARD_RATE)

D_VAL = V_VAL ^ 0.5
NORMAL_RAND_ARR = VECTOR_RANDOM_BOX_MULLER_FUNC(nLOOPS)
For i = 1 To nLOOPS
  R_VAL = E_VAL + D_VAL * NORMAL_RAND_ARR(i)
  PT_VAL = T_VAL * Exp(-BT_VAL * R_VAL)
  CALL_VAL = P0T_VAL * MAXIMUM_FUNC(PT_VAL - STRIKE, 0)
  PUT_VAL = P0T_VAL * MAXIMUM_FUNC(STRIKE - PT_VAL, 0)
  CALL_PAYOFF_SUM_VAL = CALL_PAYOFF_SUM_VAL + CALL_VAL
  PUT_PAYOFF_SUM_VAL = PUT_PAYOFF_SUM_VAL + PUT_VAL
Next i

HW_BOND_OPTION_MC_FUNC = Array(CALL_PAYOFF_SUM_VAL / nLOOPS, PUT_PAYOFF_SUM_VAL / nLOOPS)
'averagepayoffcall --> CALL_PAYOFF_SUM_VAL / nLOOPS
'averagepayoffput --> PUT_PAYOFF_SUM_VAL / nLOOPS

Exit Function
ERROR_LABEL:
HW_BOND_OPTION_MC_FUNC = Err.number
End Function

'**********************************************************************************
'**********************************************************************************
'FUNCTION      : BFN_FUNC
'DESCRIPTION   :
'LIBRARY       : FIXED_INCOME
'GROUP         : HW_SIMUL
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'**********************************************************************************
'**********************************************************************************

Private Function BFN_FUNC(ByVal HW_RATE As Double, _
ByVal T1_VAL As Double, _
ByVal T2_VAL As Double)

On Error GoTo ERROR_LABEL

BFN_FUNC = (1 - Exp(-HW_RATE * (T2_VAL - T1_VAL))) / HW_RATE

Exit Function
ERROR_LABEL:
BFN_FUNC = Err.number
End Function

'**********************************************************************************
'**********************************************************************************
'FUNCTION      : FM_FUNC
'DESCRIPTION   :
'LIBRARY       : FIXED_INCOME
'GROUP         : HW_SIMUL
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'**********************************************************************************
'**********************************************************************************

Private Function FM_FUNC(ByVal FLAT_FORWARD_RATE As Double, _
ByVal T1_VAL As Double, _
ByVal T2_VAL As Double)

On Error GoTo ERROR_LABEL

FM_FUNC = FLAT_FORWARD_RATE

Exit Function
ERROR_LABEL:
FM_FUNC = Err.number
End Function



'**********************************************************************************
'**********************************************************************************
'FUNCTION      : MARKET_BOND_PRICE_FUNC
'DESCRIPTION   :
'LIBRARY       : FIXED_INCOME
'GROUP         : HW_SIMUL
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'**********************************************************************************
'**********************************************************************************

Private Function MARKET_BOND_PRICE_FUNC(ByVal FLAT_FORWARD_RATE As Double, _
ByVal T_VAL As Double)

On Error GoTo ERROR_LABEL

MARKET_BOND_PRICE_FUNC = Exp(-FLAT_FORWARD_RATE * T_VAL)

Exit Function
ERROR_LABEL:
MARKET_BOND_PRICE_FUNC = Err.number
End Function


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : AFN_FUNC
'DESCRIPTION   :
'LIBRARY       : FIXED_INCOME
'GROUP         : HW_SIMUL
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'**********************************************************************************
'**********************************************************************************

Private Function AFN_FUNC(ByVal HW_RATE As Double, _
ByVal HW_VOLATILITY As Double, _
ByVal FLAT_FORWARD_RATE As Double, _
ByVal T1_VAL As Double, _
ByVal T2_VAL As Double)

Dim BT_VAL As Double
Dim TEMP_VAL As Double
Dim PM1_VAL As Double
Dim PM2_VAL As Double

On Error GoTo ERROR_LABEL

BT_VAL = BFN_FUNC(HW_RATE, T1_VAL, T2_VAL)
'test1 = HW_VOLATILITY * BT_VAL
'test2 = (HW_VOLATILITY * HW_VOLATILITY / (4 * HW_RATE)) * (1 - Exp(-2 * HW_RATE * T1_VAL)) * BT_VAL * BT_VAL
'test3 = test1 - test2
TEMP_VAL = BT_VAL * FM_FUNC(FLAT_FORWARD_RATE, 0, T1_VAL) - _
          (HW_VOLATILITY * HW_VOLATILITY / (4 * HW_RATE)) * _
          (1 - Exp(-2 * HW_RATE * T1_VAL)) * BT_VAL * BT_VAL
PM1_VAL = MARKET_BOND_PRICE_FUNC(FLAT_FORWARD_RATE, T1_VAL)
PM2_VAL = MARKET_BOND_PRICE_FUNC(FLAT_FORWARD_RATE, T2_VAL)
AFN_FUNC = (PM2_VAL / PM1_VAL) * Exp(TEMP_VAL)

Exit Function
ERROR_LABEL:
AFN_FUNC = Err.number
End Function


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : MFN_FUNC
'DESCRIPTION   :
'LIBRARY       : FIXED_INCOME
'GROUP         : HW_SIMUL
'ID            : 007
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'**********************************************************************************
'**********************************************************************************

Private Function MFN_FUNC(ByVal HW_RATE As Double, _
ByVal HW_VOLATILITY As Double, _
ByVal S_VAL As Double, _
ByVal T_VAL As Double, _
ByVal FORWARD_VAL As Double)

Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double

On Error GoTo ERROR_LABEL

ATEMP_VAL = (HW_VOLATILITY * HW_VOLATILITY / (HW_RATE * HW_RATE)) * _
            (1 - Exp(-HW_RATE * (T_VAL - S_VAL)))
BTEMP_VAL = (HW_VOLATILITY * HW_VOLATILITY / (2 * HW_RATE * HW_RATE)) * _
            (Exp(-HW_RATE * (FORWARD_VAL - T_VAL)) - Exp(-HW_RATE * _
            (FORWARD_VAL + T_VAL - 2 * S_VAL)))
MFN_FUNC = ATEMP_VAL - BTEMP_VAL

Exit Function
ERROR_LABEL:
MFN_FUNC = Err.number
End Function

'**********************************************************************************
'**********************************************************************************
'FUNCTION      : ALPHA_T_FUNC
'DESCRIPTION   :
'LIBRARY       : FIXED_INCOME
'GROUP         : HW_SIMUL
'ID            : 008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'**********************************************************************************
'**********************************************************************************

Private Function ALPHA_T_FUNC(ByVal HW_RATE As Double, _
ByVal HW_VOLATILITY As Double, _
ByVal FLAT_FORWARD_RATE As Double, _
ByVal T_VAL As Double)

On Error GoTo ERROR_LABEL

ALPHA_T_FUNC = FM_FUNC(FLAT_FORWARD_RATE, 0, T_VAL) + ((HW_VOLATILITY * HW_VOLATILITY) / _
                (2 * HW_RATE * HW_RATE)) * (1 - Exp(-HW_RATE * T_VAL)) ^ 2

Exit Function
ERROR_LABEL:
ALPHA_T_FUNC = Err.number
End Function


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : DISCOUNT_BOND_FUNC
'DESCRIPTION   :
'LIBRARY       : FIXED_INCOME
'GROUP         : HW_SIMUL
'ID            : 009
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'**********************************************************************************
'**********************************************************************************

Private Function DISCOUNT_BOND_FUNC(ByVal HW_RATE As Double, _
ByVal HW_VOLATILITY As Double, _
ByVal FLAT_FORWARD_RATE As Double, _
ByVal r As Double, _
ByVal T1_VAL As Double, _
ByVal T2_VAL As Double)

On Error GoTo ERROR_LABEL

DISCOUNT_BOND_FUNC = AFN_FUNC(HW_RATE, HW_VOLATILITY, FLAT_FORWARD_RATE, T1_VAL, T2_VAL) * _
                Exp(-BFN_FUNC(HW_RATE, T1_VAL, T2_VAL) * r)

Exit Function
ERROR_LABEL:
DISCOUNT_BOND_FUNC = Err.number
End Function


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : BOND_OPTION_DISCOUNT_FUNC
'DESCRIPTION   :
'LIBRARY       : FIXED_INCOME
'GROUP         : HW_SIMUL
'ID            : 010
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'**********************************************************************************
'**********************************************************************************

Private Function BOND_OPTION_DISCOUNT_FUNC(ByVal HW_RATE As Double, _
ByVal HW_VOLATILITY As Double, _
ByVal FLAT_FORWARD_RATE As Double, _
ByVal STRIKE As Double, _
ByVal T_VAL As Double, _
ByVal SWAP_START_TENOR As Double, _
ByVal MATURITY_VAL As Double, _
Optional ByVal OPTION_FLAG As Integer = 1, _
Optional ByVal CND_TYPE As Integer = 1)

Dim H_VAL As Double
Dim PS_VAL As Double
Dim PT_VAL As Double
Dim TEMP_VAL As Double
Dim SIGMA_VAL As Double

On Error GoTo ERROR_LABEL

If OPTION_FLAG <> 1 Then: OPTION_FLAG = -1
TEMP_VAL = (1 - Exp(-2 * HW_RATE * (SWAP_START_TENOR - T_VAL))) / (2 * HW_RATE)
SIGMA_VAL = HW_VOLATILITY * TEMP_VAL ^ 0.5 * BFN_FUNC(HW_RATE, SWAP_START_TENOR, MATURITY_VAL)
PS_VAL = DISCOUNT_BOND_FUNC(HW_RATE, HW_VOLATILITY, FLAT_FORWARD_RATE, FLAT_FORWARD_RATE, 0, MATURITY_VAL)
PT_VAL = DISCOUNT_BOND_FUNC(HW_RATE, HW_VOLATILITY, FLAT_FORWARD_RATE, FLAT_FORWARD_RATE, 0, SWAP_START_TENOR)
H_VAL = (1 / SIGMA_VAL) * Log(PS_VAL / (PT_VAL * STRIKE)) + SIGMA_VAL / 2

'test = CND_FUNC(H_VAL - SIGMA_VAL,CND_TYPE)

BOND_OPTION_DISCOUNT_FUNC = OPTION_FLAG * (PS_VAL * CND_FUNC(OPTION_FLAG * H_VAL, CND_TYPE) - _
                     STRIKE * PT_VAL * CND_FUNC(OPTION_FLAG * (H_VAL - SIGMA_VAL), CND_TYPE))

Exit Function
ERROR_LABEL:
BOND_OPTION_DISCOUNT_FUNC = Err.number
End Function

'**********************************************************************************
'**********************************************************************************
'FUNCTION      : FAIR_SWAP_RATE_FUNC
'DESCRIPTION   :
'LIBRARY       : FIXED_INCOME
'GROUP         : HW_SIMUL
'ID            : 011
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'**********************************************************************************
'**********************************************************************************

Private Function FAIR_SWAP_RATE_FUNC(ByVal HW_RATE As Double, _
ByVal HW_VOLATILITY As Double, _
ByVal FLAT_FORWARD_RATE As Double, _
ByVal ALPHA_VAL As Double, _
ByVal BETA_VAL As Double, _
ByVal DELTA As Double)

Dim i As Long
Dim k As Long
Dim TEMP_SUM As Double

On Error GoTo ERROR_LABEL

TEMP_SUM = 0

k = (BETA_VAL - ALPHA_VAL) / DELTA
TEMP_SUM = ALPHA_VAL
For i = 1 To k
  TEMP_SUM = TEMP_SUM + DELTA
  TEMP_SUM = TEMP_SUM + DELTA * _
             DISCOUNT_BOND_FUNC(HW_RATE, HW_VOLATILITY, FLAT_FORWARD_RATE, FLAT_FORWARD_RATE, 0, TEMP_SUM)
Next i
FAIR_SWAP_RATE_FUNC = _
    (DISCOUNT_BOND_FUNC(HW_RATE, HW_VOLATILITY, FLAT_FORWARD_RATE, FLAT_FORWARD_RATE, 0, ALPHA_VAL) - _
     DISCOUNT_BOND_FUNC(HW_RATE, HW_VOLATILITY, FLAT_FORWARD_RATE, FLAT_FORWARD_RATE, 0, BETA_VAL)) / TEMP_SUM

Exit Function
ERROR_LABEL:
FAIR_SWAP_RATE_FUNC = Err.number
End Function

