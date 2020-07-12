Attribute VB_Name = "FINAN_DERIV_BINOM_PRICING_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1      'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'Cox-Ross-Rubinstein binomial

Function OPTION_BINOMIAL_TREE1_FUNC(ByVal S_VAL As Double, _
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

'S_VAL --> Underlying Price
'K_VAL --> Strike Price
'V_VAL --> Volatility
'T_VAL --> Tenor in Years
'RF_VAL --> Risk Free Rate
'DY_VAL -->  Dividend Yield
'N_VAL --> Number of Steps
'OT_VAL --> Option Type if OT_VAL = 1 --> Call otherwise Put

Dim i As Long
Dim j As Long
Dim k As Long

Dim CC_VAL As Double 'Carrying Cost
Dim DT_VAL As Double

Dim UP_VAL As Double 'up factor
Dim DN_VAL As Double 'down factor
Dim DF_VAL As Double 'Discount
Dim PB_VAL As Double 'Pu

Dim TEMP11_VAL As Double
Dim TEMP01_VAL As Double

Dim TEMP20_VAL As Double
Dim TEMP10_VAL As Double
Dim TEMP00_VAL As Double

Dim NODES_ARR As Variant
Dim GREEKS_FLAG As Boolean

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
    OPTION_BINOMIAL_TREE1_FUNC = NODES_ARR(0)
'-----------------------------------------------------------------------------------------------------------------------------
Case 1 'Delta
'-----------------------------------------------------------------------------------------------------------------------------
    OPTION_BINOMIAL_TREE1_FUNC = (TEMP11_VAL - TEMP01_VAL) / (S_VAL * (UP_VAL - DN_VAL))
'-----------------------------------------------------------------------------------------------------------------------------
Case 2 'Gamma
'-----------------------------------------------------------------------------------------------------------------------------
    OPTION_BINOMIAL_TREE1_FUNC = (((TEMP20_VAL - TEMP10_VAL) / (S_VAL * ((UP_VAL ^ 2) - 1))) - ((TEMP10_VAL - TEMP00_VAL) / (S_VAL * (1 - (DN_VAL ^ 2))))) / (0.5 * S_VAL * ((UP_VAL ^ 2) - (DN_VAL ^ 2)))
'-----------------------------------------------------------------------------------------------------------------------------
Case 3 'x-Day Theta
'-----------------------------------------------------------------------------------------------------------------------------
'    OPTION_BINOMIAL_TREE1_FUNC = (TEMP10_VAL - NODES_ARR(0)) / (2 * DT_VAL)
    OPTION_BINOMIAL_TREE1_FUNC = OPTION_BINOMIAL_TREE1_FUNC(S_VAL, K_VAL, DTHETA_VAL, RF_VAL, DY_VAL, V_VAL, N_VAL, OT_VAL, ET_VAL, 0)
'-----------------------------------------------------------------------------------------------------------------------------
Case 4 'Vega
'-----------------------------------------------------------------------------------------------------------------------------
    OPTION_BINOMIAL_TREE1_FUNC = OPTION_BINOMIAL_TREE1_FUNC(S_VAL, K_VAL, T_VAL, RF_VAL, DY_VAL, DVEGA_VAL, N_VAL, OT_VAL, ET_VAL, 0)
'-----------------------------------------------------------------------------------------------------------------------------
Case 5 'Rho
'-----------------------------------------------------------------------------------------------------------------------------
    OPTION_BINOMIAL_TREE1_FUNC = OPTION_BINOMIAL_TREE1_FUNC(S_VAL, K_VAL, T_VAL, DRHO_VAL, DY_VAL, V_VAL, N_VAL, OT_VAL, ET_VAL, 0)
'-----------------------------------------------------------------------------------------------------------------------------
Case Else 'Implied Volatility Guess
'-----------------------------------------------------------------------------------------------------------------------------
    OPTION_BINOMIAL_TREE1_FUNC = OPTION_BINOMIAL_TREE1_FUNC(S_VAL, K_VAL, T_VAL, RF_VAL, DY_VAL, V_VAL, N_VAL, OT_VAL, ET_VAL, 4, , DVEGA_VAL, , OMP_VAL)
    OPTION_BINOMIAL_TREE1_FUNC = ((OMP_VAL * (V_VAL - DVEGA_VAL)) + (NODES_ARR(0) * DVEGA_VAL - OPTION_BINOMIAL_TREE1_FUNC * V_VAL)) / (NODES_ARR(0) - OPTION_BINOMIAL_TREE1_FUNC) 'Implied Volatility Guess
    'Implied Volatility Guess: =((Option_Market_Price*(V_1-Vega_V))+(Bin_O_1*Vega_V-Bin_O_2*V_1))/(Bin_O_1-Bin_O_2)
'-----------------------------------------------------------------------------------------------------------------------------
End Select
'-----------------------------------------------------------------------------------------------------------------------------

'-----------------------------------------------------------------------------------------------------------------------------
Exit Function
'-----------------------------------------------------------------------------------------------------------------------------
PROB_LINE:
'-----------------------------------------------------------------------------------------------------------------------------
    CC_VAL = RF_VAL - DY_VAL
    DT_VAL = T_VAL / N_VAL 'Time STEP
    UP_VAL = Exp(V_VAL * Sqr(DT_VAL)) 'Up jump - Size
    DN_VAL = 1 / UP_VAL 'Exp(-V_VAL * Sqr(DT_VAL)) - Down jump - Size
    PB_VAL = (Exp(CC_VAL * DT_VAL) - DN_VAL) / (UP_VAL - DN_VAL) 'Up probability
    DF_VAL = Exp(-RF_VAL * DT_VAL)
    'CC_VAL * Sqr(DT_VAL) --> Critical Volat -> Lower Volat than this will give negative probabilities in tree
    'Int(T_VAL / (V_VAL / CC_VAL) ^ 2) + 1 --> Minimum time steps --> Number of time steps equal or higher that will avoid negative probabilities
'-----------------------------------------------------------------------------------------------------------------------------
Return
'-----------------------------------------------------------------------------------------------------------------------------
TREE_LINE:
'-----------------------------------------------------------------------------------------------------------------------------
    ReDim NODES_ARR(0 To N_VAL + 1)
    For i = 0 To N_VAL: NODES_ARR(i) = MAXIMUM_FUNC(0, OT_VAL * (S_VAL * UP_VAL ^ i * DN_VAL ^ (N_VAL - i) - K_VAL)): Next i
    For j = N_VAL - 1 To 0 Step -1
        For i = 0 To j
            If ET_VAL = 0 Then GoSub EURO_LINE Else GoSub AMER_LINE
            If ((GREEKS_FLAG = True) And ((j = 2) Or (j = 1))) Then: GoSub GREEKS_LINE
        Next i
    Next j
'-----------------------------------------------------------------------------------------------------------------------------
Return
'-----------------------------------------------------------------------------------------------------------------------------
EURO_LINE:
'-----------------------------------------------------------------------------------------------------------------------------
    NODES_ARR(i) = (PB_VAL * NODES_ARR(i + 1) + (1 - PB_VAL) * NODES_ARR(i)) * DF_VAL
'-----------------------------------------------------------------------------------------------------------------------------
Return
'-----------------------------------------------------------------------------------------------------------------------------
AMER_LINE:
'-----------------------------------------------------------------------------------------------------------------------------
    NODES_ARR(i) = MAXIMUM_FUNC((OT_VAL * (S_VAL * UP_VAL ^ i * DN_VAL ^ (j - i) - K_VAL)), (PB_VAL * NODES_ARR(i + 1) + (1 - PB_VAL) * NODES_ARR(i)) * DF_VAL)
'-----------------------------------------------------------------------------------------------------------------------------
Return
'-----------------------------------------------------------------------------------------------------------------------------
GREEKS_LINE:
'-----------------------------------------------------------------------------------------------------------------------------
    'Debug.Print NODES_ARR(i), i, j
    Select Case j
    Case 1 'Delta
        TEMP11_VAL = NODES_ARR(1)
        TEMP01_VAL = NODES_ARR(0)
    Case 2 'Gamma
        TEMP20_VAL = NODES_ARR(2)
        TEMP10_VAL = NODES_ARR(1) 'Theta
        TEMP00_VAL = NODES_ARR(0)
    End Select
'-----------------------------------------------------------------------------------------------------------------------------
Return
'-----------------------------------------------------------------------------------------------------------------------------
ERROR_LABEL:
OPTION_BINOMIAL_TREE1_FUNC = Err.number
End Function


Function OPTION_BINOMIAL_TREE2_FUNC(ByVal S_VAL As Double, _
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
Dim PB_VAL As Double 'Pu

Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

If OT_VAL = 1 Then k = 1 Else k = -1
'For Vega just adjust the Volalitity (V_VAL)
'For Rho just adjust the Riskfree (RF_VAL)
GoSub PROB_LINE
GoSub TREE_LINE

'-----------------------------------------------------------------------------------------------------------------------------
Select Case OUTPUT
'-----------------------------------------------------------------------------------------------------------------------------
Case 0 'Option Fair Price
'-----------------------------------------------------------------------------------------------------------------------------
    OPTION_BINOMIAL_TREE2_FUNC = TEMP_MATRIX(N_VAL + 1, N_VAL + 2)
'-----------------------------------------------------------------------------------------------------------------------------
Case 1 'Delta
'-----------------------------------------------------------------------------------------------------------------------------
    OPTION_BINOMIAL_TREE2_FUNC = (TEMP_MATRIX(N_VAL + 0, N_VAL + 1) - TEMP_MATRIX(N_VAL + 1, N_VAL + 1)) / (S_VAL * (UP_VAL - DN_VAL))
'-----------------------------------------------------------------------------------------------------------------------------
Case 2 'Gamma
'-----------------------------------------------------------------------------------------------------------------------------
    OPTION_BINOMIAL_TREE2_FUNC = (((TEMP_MATRIX(N_VAL - 1, N_VAL + 0) - TEMP_MATRIX(N_VAL + 0, N_VAL + 0)) / (S_VAL * ((UP_VAL ^ 2) - 1))) - ((TEMP_MATRIX(N_VAL + 0, N_VAL + 0) - TEMP_MATRIX(N_VAL + 1, N_VAL + 0)) / (S_VAL * (1 - (DN_VAL ^ 2))))) / (0.5 * S_VAL * ((UP_VAL ^ 2) - (DN_VAL ^ 2)))
'-----------------------------------------------------------------------------------------------------------------------------
Case 3 'x-Day Theta
'-----------------------------------------------------------------------------------------------------------------------------
    OPTION_BINOMIAL_TREE2_FUNC = OPTION_BINOMIAL_TREE2_FUNC(S_VAL, K_VAL, DTHETA_VAL, RF_VAL, DY_VAL, V_VAL, N_VAL, OT_VAL, 0)
'-----------------------------------------------------------------------------------------------------------------------------
Case 4 'Vega
'-----------------------------------------------------------------------------------------------------------------------------
    OPTION_BINOMIAL_TREE2_FUNC = OPTION_BINOMIAL_TREE2_FUNC(S_VAL, K_VAL, T_VAL, RF_VAL, DY_VAL, DVEGA_VAL, N_VAL, OT_VAL, 0)
'-----------------------------------------------------------------------------------------------------------------------------
Case 5 'Rho
'-----------------------------------------------------------------------------------------------------------------------------
    OPTION_BINOMIAL_TREE2_FUNC = OPTION_BINOMIAL_TREE2_FUNC(S_VAL, K_VAL, T_VAL, DRHO_VAL, DY_VAL, V_VAL, N_VAL, OT_VAL, 0)
'-----------------------------------------------------------------------------------------------------------------------------
Case Else 'Implied Volatility Guess
'-----------------------------------------------------------------------------------------------------------------------------
    OPTION_BINOMIAL_TREE2_FUNC = OPTION_BINOMIAL_TREE2_FUNC(S_VAL, K_VAL, T_VAL, RF_VAL, DY_VAL, V_VAL, N_VAL, OT_VAL, 4, , DVEGA_VAL, , OMP_VAL)
    OPTION_BINOMIAL_TREE2_FUNC = ((OMP_VAL * (V_VAL - DVEGA_VAL)) + (TEMP_MATRIX(N_VAL + 1, N_VAL + 2) * DVEGA_VAL - OPTION_BINOMIAL_TREE2_FUNC * V_VAL)) / (TEMP_MATRIX(N_VAL + 1, N_VAL + 2) - OPTION_BINOMIAL_TREE2_FUNC) 'Implied Volatility Guess
    'Implied Volatility Guess: =((Option_Market_Price*(V_1-Vega_V))+(Bin_O_1*Vega_V-Bin_O_2*V_1))/(Bin_O_1-Bin_O_2)
'-----------------------------------------------------------------------------------------------------------------------------
End Select
'-----------------------------------------------------------------------------------------------------------------------------

'-----------------------------------------------------------------------------------------------------------------------------
Exit Function
'-----------------------------------------------------------------------------------------------------------------------------
PROB_LINE:
'-----------------------------------------------------------------------------------------------------------------------------
    If DY_VAL <> 0 Then AS_VAL = S_VAL * Exp(-DY_VAL * T_VAL) Else AS_VAL = S_VAL '--> Also for Theta
    DT_VAL = T_VAL / N_VAL
    UP_VAL = Exp(V_VAL * Sqr(DT_VAL)) 'up factor
    DN_VAL = 1 / UP_VAL 'down factor
    DF_VAL = Exp(-RF_VAL * DT_VAL)
    PB_VAL = (Exp(RF_VAL * DT_VAL) - DN_VAL) / (UP_VAL - DN_VAL) 'Prob

'-----------------------------------------------------------------------------------------------------------------------------
Return
'-----------------------------------------------------------------------------------------------------------------------------
TREE_LINE:
'-----------------------------------------------------------------------------------------------------------------------------
    l = N_VAL + 1
    ReDim TEMP_MATRIX(1 To l, 1 To N_VAL + 2)
    For i = 1 To l
        h = N_VAL - i + 1
        TEMP_MATRIX(i, 1) = h
        L_VAL = k * (-K_VAL + (UP_VAL ^ h) * (DN_VAL ^ (N_VAL - h)) * AS_VAL)
        TEMP_MATRIX(i, 2) = IIf(L_VAL > 0, L_VAL, 0)
    Next i
    'Option Fair Price
    For j = 1 To N_VAL
        For i = j + 1 To l
            TEMP_MATRIX(i, 2 + j) = ((PB_VAL * TEMP_MATRIX(i - 1, 1 + j)) + ((1 - PB_VAL) * TEMP_MATRIX(i - 0, 1 + j))) * DF_VAL
        Next i
    Next j
'-----------------------------------------------------------------------------------------------------------------------------
Return
'-----------------------------------------------------------------------------------------------------------------------------
ERROR_LABEL:
OPTION_BINOMIAL_TREE2_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CRR_BINOMIAL_TREE_MD_FUNC
'DESCRIPTION   : Cox-Ross-Rubinstein binomial tree multidimensional
'LIBRARY       : DERIVATIVES
'GROUP         : BINOMIAL
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function CRR_BINOMIAL_TREE_MD_FUNC( _
ByVal S_VAL As Double, _
ByVal K_VAL As Double, _
ByVal T_VAL As Double, _
ByVal RF_VAL As Double, _
ByVal CC_VAL As Double, _
ByVal V_VAL As Double, _
Optional ByVal N_VAL As Long = 150, _
Optional ByVal OT_VAL As Integer = 1, _
Optional ByVal ET_VAL As Integer = 0, _
Optional ByVal MODEL As Integer = 0)

'CC_VAL = RF_VAL - DY_VAL
'ET_VAL = 0 for European, Else for American

Dim i As Long
Dim j As Long
Dim k As Long

Dim UP_VAL As Double
Dim DN_VAL As Double
Dim PB_VAL As Double
Dim DT_VAL As Double
Dim DF_VAL As Double

Dim NODE_ARR As Variant
Dim TREE_MATRIX As Variant

On Error GoTo ERROR_LABEL

'-------------------------------------------------------------------------------------
Select Case MODEL
'-------------------------------------------------------------------------------------
Case 0 'The well-known Cox-Ross-Rubinstein binomial tree can be used to price
' European and American options
' - on stocks w/o dividends (b=r)
' - on stocks w/o paying continuous dividend yield q (b=r-q)
' - futures (b=0)
' - on currencies (CC_VAL = RF_VAL - rforeign)
'-------------------------------------------------------------------------------------
    ReDim NODE_ARR(0 To N_VAL + 1)

    If OT_VAL = 1 Then k = 1 Else k = -1
    DT_VAL = T_VAL / N_VAL
    UP_VAL = Exp(V_VAL * Sqr(DT_VAL))
    DN_VAL = 1 / UP_VAL
    PB_VAL = (Exp(CC_VAL * DT_VAL) - DN_VAL) / (UP_VAL - DN_VAL)
    DF_VAL = Exp(-RF_VAL * DT_VAL)
    
    For i = 0 To N_VAL
        NODE_ARR(i) = MAXIMUM_FUNC(0, k * (S_VAL * UP_VAL ^ i * DN_VAL ^ (N_VAL - i) - K_VAL))
    Next i
    
    For j = N_VAL - 1 To 0 Step -1:
        For i = 0 To j
            If ET_VAL = 0 Then 'European
                NODE_ARR(i) = (PB_VAL * NODE_ARR(i + 1) + (1 - PB_VAL) * NODE_ARR(i)) * DF_VAL
            Else 'If ET_VAL = "a" Then 'American
                NODE_ARR(i) = MAXIMUM_FUNC((k * (S_VAL * UP_VAL ^ i * DN_VAL ^ (Abs(i - j)) - K_VAL)), (PB_VAL * NODE_ARR(i + 1) + (1 - PB_VAL) * NODE_ARR(i)) * DF_VAL)
            End If
        Next i
    Next j

    CRR_BINOMIAL_TREE_MD_FUNC = NODE_ARR(0) 'default first index of arrays.
'-------------------------------------------------------------------------------------
Case Else
' This adapted 3D function returns at each node of the binomial tree
'  - the value of the underlying asset (S_VAL) - NODE_ARR(i,j,1)
'  - option price - NODE_ARR(i,j,2)'
'-------------------------------------------------------------------------------------
    ReDim NODE_ARR(0 To N_VAL + 1, 0 To N_VAL, 0 To 1)
    ReDim TREE_MATRIX(0 To 2 * (N_VAL + 1) - 1, 0 To N_VAL)
    If OT_VAL = 1 Then k = 1 Else k = -1
    
    DT_VAL = T_VAL / N_VAL
    UP_VAL = Exp(V_VAL * Sqr(DT_VAL))
    DN_VAL = 1 / UP_VAL
    PB_VAL = (Exp(CC_VAL * DT_VAL) - DN_VAL) / (UP_VAL - DN_VAL)
    DF_VAL = Exp(-RF_VAL * DT_VAL)
    
    For i = 0 To N_VAL
        NODE_ARR(i, N_VAL, 0) = S_VAL * UP_VAL ^ i * DN_VAL ^ (N_VAL - i)
        NODE_ARR(i, N_VAL, 1) = MAXIMUM_FUNC(0, k * (S_VAL * UP_VAL ^ i * DN_VAL ^ (N_VAL - i) - K_VAL))
    Next i
    
    For j = N_VAL - 1 To 0 Step -1:
        For i = 0 To j
            NODE_ARR(i, j, 0) = S_VAL * UP_VAL ^ i * DN_VAL ^ (j - i)
            If ET_VAL = 0 Then 'European
                NODE_ARR(i, j, 1) = (PB_VAL * NODE_ARR(i + 1, j + 1, 1) + (1 - PB_VAL) * NODE_ARR(i, j + 1, 1)) * DF_VAL
            Else 'If ET_VAL = "a" Then 'American
                NODE_ARR(i, j, 1) = MAXIMUM_FUNC((k * (S_VAL * UP_VAL ^ i * DN_VAL ^ (Abs(i - j)) - K_VAL)), (PB_VAL * NODE_ARR(i + 1, j + 1, 1) + (1 - PB_VAL) * NODE_ARR(i, j + 1, 1)) * DF_VAL)
            End If
        Next i
    Next j
      'initialize TREE_MATRIX()
    For j = 0 To UBound(TREE_MATRIX, 2)
       For i = 0 To UBound(TREE_MATRIX, 1)
            TREE_MATRIX(i, j) = " "
       Next i
    Next j
    For j = N_VAL To 0 Step -1   'columns
       For i = 0 To j 'rows
            TREE_MATRIX(2 * (j - i), j) = NODE_ARR(i, j, 0)
            TREE_MATRIX(2 * (j - i) + 1, j) = NODE_ARR(i, j, 1)
       Next i
    Next j
    CRR_BINOMIAL_TREE_MD_FUNC = TREE_MATRIX
'-------------------------------------------------------------------------------------
End Select
'-------------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
CRR_BINOMIAL_TREE_MD_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CRR_BINOMIAL_TREE_AMER_FUNC
'DESCRIPTION   : Classical binomial model
'    Simplifying algebraic terms in the binomial model of
'    Cox-Ross-Rubinstein for american options the
'    speed against the usual solution is improved by a
'    factor of 40 - 50.
'----------------------------------------------------------------------------
'    It does not heal the combinatorical curse of double
'    looping for American options ...
'----------------------------------------------------------------------------
'LIBRARY       : DERIVATIVES
'GROUP         : BINOMIAL
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function CRR_BINOMIAL_TREE_AMER_FUNC(ByVal S_VAL As Double, _
ByVal K_VAL As Double, _
ByVal T_VAL As Double, _
ByVal RF_VAL As Double, _
ByVal DY_VAL As Double, _
ByVal V_VAL As Double, _
Optional ByVal N_VAL As Long = 150, _
Optional ByVal OT_VAL As Integer = 1)

Dim i As Long
Dim j As Long

Dim ii As Double
Dim jj As Double

Dim DT_VAL As Double
Dim MT_VAL As Double
Dim FT_VAL As Double
Dim IFT_VAL As Double

Dim X_VAL As Double
Dim UP_VAL As Double
Dim DN_VAL As Double
Dim PB_VAL As Double

Dim SR_VAL As Double
Dim SQ_VAL As Double
Dim TM_VAL As Double

Dim NODES_ARR As Variant ' working array for Leisen Reimer tree

On Error GoTo ERROR_LABEL

ReDim NODES_ARR(0 To N_VAL + 1)
If OT_VAL <> 1 Then: OT_VAL = -1
If (0 < N_VAL And 0 < S_VAL And 0 < K_VAL And 0 < T_VAL And 0 < V_VAL) Then
Else
    CRR_BINOMIAL_TREE_AMER_FUNC = -2
    Exit Function
End If

X_VAL = K_VAL / S_VAL

DT_VAL = T_VAL / CDbl(N_VAL)
FT_VAL = Exp(RF_VAL * DT_VAL)
MT_VAL = Exp((RF_VAL - DY_VAL) * DT_VAL)

UP_VAL = Exp(V_VAL * Sqr(DT_VAL))
DN_VAL = 1 / UP_VAL
PB_VAL = (MT_VAL - DN_VAL) / (UP_VAL - DN_VAL)
SR_VAL = 1 - PB_VAL
IFT_VAL = 1 / FT_VAL

SQ_VAL = UP_VAL * UP_VAL

ii = UP_VAL ^ (-2 - N_VAL)
For i = 1 To N_VAL + 1
    ii = SQ_VAL * ii
    TM_VAL = OT_VAL * (ii - X_VAL) ' check against pay off
    If (0 <= TM_VAL) Then
        NODES_ARR(i - 1) = TM_VAL
    Else
        NODES_ARR(i - 1) = 0#
    End If
Next i

jj = UP_VAL ^ (-2 - N_VAL)
For j = N_VAL To 1 Step -1
    jj = jj * UP_VAL
    ii = jj
    For i = 1 To j
        NODES_ARR(i - 1) = (PB_VAL * NODES_ARR(i) + SR_VAL * NODES_ARR(i - 1)) * IFT_VAL
        ii = SQ_VAL * ii
        TM_VAL = OT_VAL * (ii - X_VAL) ' check against pay off
        If (NODES_ARR(i - 1) <= TM_VAL) Then: NODES_ARR(i - 1) = TM_VAL
    Next i
Next j

CRR_BINOMIAL_TREE_AMER_FUNC = S_VAL * NODES_ARR(0)
      
Exit Function
ERROR_LABEL:
CRR_BINOMIAL_TREE_AMER_FUNC = Err.number
End Function

