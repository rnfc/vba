Attribute VB_Name = "FINAN_DERIV_BINOM_TRI_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1      'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'The following model is built under the assumption that cash paid to
'owners of the underlying, such as dividends and interest, are paid
'continuously at constant rate over the life of the option. This
'assumption is relatively accurate for valuing puts generally, and calls
'on bonds, commodities, currencies and stock index portfolios.

Function OPTION_TRINOMIAL_TREE1_FUNC(ByVal S_VAL As Double, _
ByVal K_VAL As Double, _
ByVal T_VAL As Double, _
ByVal RF_VAL As Double, _
ByVal DY_VAL As Double, _
ByVal V_VAL As Double, _
Optional ByVal N_VAL As Long = 150, _
Optional ByVal OT_VAL As Integer = 1, _
Optional ByVal ET_VAL As Integer = 0, _
Optional ByVal OUTPUT As Integer = 0, _
Optional ByVal DTHETA_VAL As Double = 0, _
Optional ByVal DVEGA_VAL As Double = 0, _
Optional ByVal DRHO_VAL As Double = 0, _
Optional ByVal OMP_VAL As Double = 0)

'S_VAL: Underlying asset current price.

'K_VAL: Price at which the underlying can be bought (for calls) or
'sold (for puts).

'T_VAL: Maturity scaled on an annualized basis

'RF_VAL: The cost of funds for the number of days to MATURITY
'It should corresponds to the period to the record date, not the option MATURITY

'DY_VAL: Yield paid on the underlying asset matching
'the option's MATURITY.

'V_VAL: Annualized volatility of the underlying asset price

'N_VAL --> Number of N_VAL (e.g. One-month intervals)

'OT_VAL: 1 for call, and -1 for put.
'ET_VAL: 0 For European, else For American


'DTHETA_VAL --> Days
'DVEGA_VAL --> Volatility
'DRHO_VAL --> Rates
'OMP_VAL --> Option Market Price

Dim i As Long
Dim j As Long
Dim k As Long

Dim DT_VAL As Double

Dim UP_VAL As Double 'up factor
Dim DN_VAL As Double 'down factor
Dim DF_VAL As Double 'Discount
Dim PU_VAL As Double 'Pu
Dim PD_VAL As Double 'Pd
Dim PM_VAL As Double 'Pm

Dim GREEKS_FLAG As Boolean
Dim TEMP02_VAL As Double
Dim TEMP22_VAL As Double
Dim TEMP42_VAL As Double
Dim TEMP01_VAL As Double
Dim TEMP11_VAL As Double
Dim TEMP21_VAL As Double
Dim TEMP00_VAL As Double

Dim NODES_ARR As Variant

On Error GoTo ERROR_LABEL

If OT_VAL = 1 Then k = 1 Else k = -1
If OUTPUT > 0 And OUTPUT < 6 Then: GREEKS_FLAG = True
GoSub PROB_LINE
GoSub TREE_LINE

'-----------------------------------------------------------------------------------------------------------------------------
Select Case OUTPUT
'-----------------------------------------------------------------------------------------------------------------------------
Case 0 'Option Fair Price
'-----------------------------------------------------------------------------------------------------------------------------
    OPTION_TRINOMIAL_TREE1_FUNC = NODES_ARR(0)
'-----------------------------------------------------------------------------------------------------------------------------
Case 1 'Delta
'-----------------------------------------------------------------------------------------------------------------------------
    OPTION_TRINOMIAL_TREE1_FUNC = (TEMP21_VAL - TEMP01_VAL) / (S_VAL * (UP_VAL - DN_VAL))
'-----------------------------------------------------------------------------------------------------------------------------
Case 2 'Gamma
'-----------------------------------------------------------------------------------------------------------------------------
    OPTION_TRINOMIAL_TREE1_FUNC = (((TEMP42_VAL - TEMP22_VAL) / (S_VAL * ((UP_VAL ^ 2) - 1))) - ((TEMP22_VAL - TEMP02_VAL) / (S_VAL * (1 - (DN_VAL ^ 2))))) / (0.5 * S_VAL * ((UP_VAL ^ 2) - (DN_VAL ^ 2)))
'-----------------------------------------------------------------------------------------------------------------------------
Case 3 'x-Day Theta
'-----------------------------------------------------------------------------------------------------------------------------
    OPTION_TRINOMIAL_TREE1_FUNC = (TEMP11_VAL - TEMP00_VAL) / (DTHETA_VAL / N_VAL)
'-----------------------------------------------------------------------------------------------------------------------------
Case 4 'Vega
'-----------------------------------------------------------------------------------------------------------------------------
    OPTION_TRINOMIAL_TREE1_FUNC = OPTION_TRINOMIAL_TREE1_FUNC(S_VAL, K_VAL, T_VAL, RF_VAL, DY_VAL, DVEGA_VAL, N_VAL, OT_VAL, ET_VAL, 0)
'-----------------------------------------------------------------------------------------------------------------------------
Case 5 'Rho
'-----------------------------------------------------------------------------------------------------------------------------
    OPTION_TRINOMIAL_TREE1_FUNC = OPTION_TRINOMIAL_TREE1_FUNC(S_VAL, K_VAL, T_VAL, DRHO_VAL, DY_VAL, V_VAL, N_VAL, OT_VAL, ET_VAL, 0)
Case Else 'Implied Volatility Guess
    OPTION_TRINOMIAL_TREE1_FUNC = OPTION_TRINOMIAL_TREE1_FUNC(S_VAL, K_VAL, T_VAL, RF_VAL, DY_VAL, V_VAL, N_VAL, OT_VAL, ET_VAL, 4, DTHETA_VAL, DVEGA_VAL, DRHO_VAL, OMP_VAL)
    OPTION_TRINOMIAL_TREE1_FUNC = ((OMP_VAL * (V_VAL - DVEGA_VAL)) + (NODES_ARR(0) * DVEGA_VAL - OPTION_TRINOMIAL_TREE1_FUNC * V_VAL)) / (NODES_ARR(0) - OPTION_TRINOMIAL_TREE1_FUNC) 'Implied Volatility Guess
    'Implied Volatility Guess: =((Option_Market_Price*(V_1-Tri_Vega_V))+(Tri_O_1*Tri_Vega_V-Tri_O_2*V_1))/(Tri_O_1-Tri_O_2)

'-----------------------------------------------------------------------------------------------------------------------------
End Select
'-----------------------------------------------------------------------------------------------------------------------------

'-----------------------------------------------------------------------------------------------------------------------------
Exit Function
'-----------------------------------------------------------------------------------------------------------------------------
PROB_LINE:
'-----------------------------------------------------------------------------------------------------------------------------
    DT_VAL = T_VAL / N_VAL
    DF_VAL = Exp(-RF_VAL * DT_VAL)
    
    UP_VAL = Exp(V_VAL * Sqr(2 * DT_VAL))
    DN_VAL = Exp(-V_VAL * Sqr(2 * DT_VAL))
    
    PU_VAL = ((Exp((RF_VAL - DY_VAL) * DT_VAL / 2) - Exp(-V_VAL * Sqr(DT_VAL / 2))) / (Exp(V_VAL * Sqr(DT_VAL / 2)) - Exp(-V_VAL * Sqr(DT_VAL / 2)))) ^ 2
    PD_VAL = ((Exp(V_VAL * Sqr(DT_VAL / 2)) - Exp((RF_VAL - DY_VAL) * DT_VAL / 2)) / (Exp(V_VAL * Sqr(DT_VAL / 2)) - Exp(-V_VAL * Sqr(DT_VAL / 2)))) ^ 2
    PM_VAL = 1 - PU_VAL - PD_VAL

'-----------------------------------------------------------------------------------------------------------------------------
Return
'-----------------------------------------------------------------------------------------------------------------------------
TREE_LINE:
'-----------------------------------------------------------------------------------------------------------------------------
    ReDim NODES_ARR(0 To 2 * N_VAL)
    For i = 0 To (2 * N_VAL): NODES_ARR(i) = MAXIMUM_FUNC(0, OT_VAL * (S_VAL * UP_VAL ^ MAXIMUM_FUNC(i - N_VAL, 0) * DN_VAL ^ MAXIMUM_FUNC(N_VAL * 2 - N_VAL - i, 0) - K_VAL)): Next i
    For j = N_VAL - 1 To 0 Step -1
        k = j * 2
        For i = 0 To k
            If ET_VAL = 0 Then GoSub EURO_LINE Else GoSub AMER_LINE
            If ((GREEKS_FLAG = True) And (j < 3)) Then: GoSub GREEKS_LINE
        Next i
    Next j
'-----------------------------------------------------------------------------------------------------------------------------
Return
'-----------------------------------------------------------------------------------------------------------------------------
EURO_LINE:
'-----------------------------------------------------------------------------------------------------------------------------
    NODES_ARR(i) = (PU_VAL * NODES_ARR(i + 2) + PM_VAL * NODES_ARR(i + 1) + PD_VAL * NODES_ARR(i)) * DF_VAL
'-----------------------------------------------------------------------------------------------------------------------------
Return
'-----------------------------------------------------------------------------------------------------------------------------
AMER_LINE:
'-----------------------------------------------------------------------------------------------------------------------------
    NODES_ARR(i) = MAXIMUM_FUNC((OT_VAL * (S_VAL * UP_VAL ^ MAXIMUM_FUNC(i - j, 0) * DN_VAL ^ MAXIMUM_FUNC(j * 2 - j - i, 0) - K_VAL)), (PU_VAL * NODES_ARR(i + 2) + PM_VAL * NODES_ARR(i + 1) + PD_VAL * NODES_ARR(i)) * DF_VAL)
'-----------------------------------------------------------------------------------------------------------------------------
Return
'-----------------------------------------------------------------------------------------------------------------------------
GREEKS_LINE:
'-----------------------------------------------------------------------------------------------------------------------------
    'Debug.Print NODES_ARR(i), i, j
    If i = 0 And j = 2 Then
        TEMP02_VAL = NODES_ARR(i)
    ElseIf i = 2 And j = 2 Then
        TEMP22_VAL = NODES_ARR(i)
    ElseIf i = 4 And j = 2 Then
        TEMP42_VAL = NODES_ARR(i)
    ElseIf i = 0 And j = 1 Then
        TEMP01_VAL = NODES_ARR(i)
    ElseIf i = 1 And j = 1 Then
        TEMP11_VAL = NODES_ARR(i)
    ElseIf i = 2 And j = 1 Then
        TEMP21_VAL = NODES_ARR(i)
    ElseIf i = 0 And j = 0 Then
        TEMP00_VAL = NODES_ARR(i)
    End If
'-----------------------------------------------------------------------------------------------------------------------------
Return
'-----------------------------------------------------------------------------------------------------------------------------
ERROR_LABEL:
OPTION_TRINOMIAL_TREE1_FUNC = Err.number
End Function

'European Only!!!
Function OPTION_TRINOMIAL_TREE2_FUNC(ByVal S_VAL As Double, _
ByVal K_VAL As Double, _
ByVal T_VAL As Double, _
ByVal RF_VAL As Double, _
ByVal DY_VAL As Double, _
ByVal V_VAL As Double, _
Optional ByVal N_VAL As Long = 150, _
Optional ByVal OT_VAL As Integer = 1, _
Optional ByVal OUTPUT As Integer = 0, _
Optional ByVal DTHETA_VAL As Double = 0, _
Optional ByVal DVEGA_VAL As Double = 0, _
Optional ByVal DRHO_VAL As Double = 0, _
Optional ByVal OMP_VAL As Double = 0)

'S_VAL --> Underlying Price
'K_VAL --> Strike Price
'V_VAL --> Volatility
'T_VAL --> Tenor in Years
'RF_VAL --> Risk Free Rate
'DY_VAL -->  Dividend Yield
'N_VAL --> Number of Steps
'OT_VAL --> Option Type if OT_VAL = 1 --> Call otherwise Put

'DTHETA_VAL --> Days
'DVEGA_VAL --> Volatility
'DRHO_VAL --> Rates
'OMP_VAL --> Option Market Price

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim DT_VAL As Double

Dim L_VAL As Double
Dim AS_VAL As Double 'Adjusted Stock Price
Dim UP_VAL As Double 'up factor
Dim DN_VAL As Double 'down factor
Dim DF_VAL As Double 'Discount
Dim PU_VAL As Double 'Pu
Dim PD_VAL As Double 'Pd
Dim PM_VAL As Double 'Pm

Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

If OT_VAL = 1 Then k = 1 Else k = -1
GoSub PROB_LINE
GoSub TREE_LINE

'-----------------------------------------------------------------------------------------------------------------------------
Select Case OUTPUT
'-----------------------------------------------------------------------------------------------------------------------------
Case 0 'Option Fair Price
'-----------------------------------------------------------------------------------------------------------------------------
    OPTION_TRINOMIAL_TREE2_FUNC = TEMP_MATRIX(N_VAL + 1, N_VAL + 2)
'-----------------------------------------------------------------------------------------------------------------------------
Case 1 'Delta
'-----------------------------------------------------------------------------------------------------------------------------
    OPTION_TRINOMIAL_TREE2_FUNC = (TEMP_MATRIX(N_VAL + 0, N_VAL + 1) - TEMP_MATRIX(N_VAL + 2, N_VAL + 1)) / (S_VAL * (UP_VAL - DN_VAL))
'-----------------------------------------------------------------------------------------------------------------------------
Case 2 'Gamma
'-----------------------------------------------------------------------------------------------------------------------------
    OPTION_TRINOMIAL_TREE2_FUNC = (((TEMP_MATRIX(N_VAL - 1, N_VAL + 0) - TEMP_MATRIX(N_VAL + 1, N_VAL + 0)) / (S_VAL * ((UP_VAL ^ 2) - 1))) - ((TEMP_MATRIX(N_VAL + 1, N_VAL + 0) - TEMP_MATRIX(N_VAL + 3, N_VAL + 0)) / (S_VAL * (1 - (DN_VAL ^ 2))))) / (0.5 * S_VAL * ((UP_VAL ^ 2) - (DN_VAL ^ 2)))
'-----------------------------------------------------------------------------------------------------------------------------
Case 3 'x-Day Theta
'-----------------------------------------------------------------------------------------------------------------------------
    OPTION_TRINOMIAL_TREE2_FUNC = (TEMP_MATRIX(N_VAL + 1, N_VAL + 1) - TEMP_MATRIX(N_VAL + 1, N_VAL + 2)) / (DTHETA_VAL / N_VAL)
'-----------------------------------------------------------------------------------------------------------------------------
Case 4 'Vega
'-----------------------------------------------------------------------------------------------------------------------------
    OPTION_TRINOMIAL_TREE2_FUNC = OPTION_TRINOMIAL_TREE2_FUNC(S_VAL, K_VAL, T_VAL, RF_VAL, DY_VAL, DVEGA_VAL, N_VAL, OT_VAL, 0)
'-----------------------------------------------------------------------------------------------------------------------------
Case 5 'Rho
'-----------------------------------------------------------------------------------------------------------------------------
    OPTION_TRINOMIAL_TREE2_FUNC = OPTION_TRINOMIAL_TREE2_FUNC(S_VAL, K_VAL, T_VAL, DRHO_VAL, DY_VAL, V_VAL, N_VAL, OT_VAL, 0)
Case Else 'Implied Volatility Guess
    OPTION_TRINOMIAL_TREE2_FUNC = OPTION_TRINOMIAL_TREE2_FUNC(S_VAL, K_VAL, T_VAL, RF_VAL, DY_VAL, V_VAL, N_VAL, OT_VAL, 4, DTHETA_VAL, DVEGA_VAL, DRHO_VAL, OMP_VAL)
    OPTION_TRINOMIAL_TREE2_FUNC = ((OMP_VAL * (V_VAL - DVEGA_VAL)) + (TEMP_MATRIX(N_VAL + 1, N_VAL + 2) * DVEGA_VAL - OPTION_TRINOMIAL_TREE2_FUNC * V_VAL)) / (TEMP_MATRIX(N_VAL + 1, N_VAL + 2) - OPTION_TRINOMIAL_TREE2_FUNC) 'Implied Volatility Guess
    'Implied Volatility Guess: =((Option_Market_Price*(V_1-Tri_Vega_V))+(Tri_O_1*Tri_Vega_V-Tri_O_2*V_1))/(Tri_O_1-Tri_O_2)

'-----------------------------------------------------------------------------------------------------------------------------
End Select
'-----------------------------------------------------------------------------------------------------------------------------

'-----------------------------------------------------------------------------------------------------------------------------
Exit Function
'-----------------------------------------------------------------------------------------------------------------------------
PROB_LINE:
'-----------------------------------------------------------------------------------------------------------------------------
    If DY_VAL <> 0 Then AS_VAL = S_VAL * Exp(-DY_VAL * T_VAL) Else AS_VAL = S_VAL
    DT_VAL = T_VAL / N_VAL
    UP_VAL = Exp(V_VAL * Sqr(2 * DT_VAL)) 'up factor
    DN_VAL = 1 / UP_VAL 'down factor
    DF_VAL = Exp(-RF_VAL * DT_VAL)
    PU_VAL = ((Exp(RF_VAL * DT_VAL / 2) - Exp(-V_VAL * Sqr(DT_VAL / 2))) / (Exp(V_VAL * Sqr(DT_VAL / 2)) - Exp(-V_VAL * Sqr(DT_VAL / 2)))) ^ 2 'Pu
    PD_VAL = ((Exp(V_VAL * Sqr(DT_VAL / 2)) - Exp(RF_VAL * DT_VAL / 2)) / (Exp(V_VAL * Sqr(DT_VAL / 2)) - Exp(-V_VAL * Sqr(DT_VAL / 2)))) ^ 2 'Pd
    PM_VAL = 1 - PU_VAL - PD_VAL 'Pm

'-----------------------------------------------------------------------------------------------------------------------------
Return
'-----------------------------------------------------------------------------------------------------------------------------
TREE_LINE:
'-----------------------------------------------------------------------------------------------------------------------------
    l = N_VAL * 2 + 1
    ReDim TEMP_MATRIX(1 To l, 1 To N_VAL + 2)
    For i = 1 To l
        h = N_VAL - i + 1
        TEMP_MATRIX(i, 1) = h
        L_VAL = k * ((UP_VAL ^ h) * AS_VAL - K_VAL)
        TEMP_MATRIX(i, 2) = IIf(L_VAL > 0, L_VAL, 0)
    Next i
    
    'Option Fair Price: TEMP_MATRIX(N_VAL + 1, N_VAL + 2)
    For j = 1 To N_VAL
'        For i = 1 To l
        For i = j + 1 To l - j
            'If i > j And i <= (l - j) Then
                TEMP_MATRIX(i, 2 + j) = DF_VAL * ((PU_VAL * TEMP_MATRIX(i - 1, 1 + j)) + (PM_VAL * TEMP_MATRIX(i, 1 + j)) + (PD_VAL * TEMP_MATRIX(i + 1, 1 + j)))
'            Else
 '               TEMP_MATRIX(i, 2 + j) = ""
  '          End If
        Next i
    Next j
'-----------------------------------------------------------------------------------------------------------------------------
Return
'-----------------------------------------------------------------------------------------------------------------------------
ERROR_LABEL:
OPTION_TRINOMIAL_TREE2_FUNC = Err.number
End Function

