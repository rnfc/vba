Attribute VB_Name = "FINAN_DERIV_BINOM_IBT_LIBR"

'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------
'The Generalized Reduced Gradient Algorithm needed to generate implied ending
'risk-neutral probabilities from a set of actual option prices and
'the backwards recursion needed to solve for the entire implied tree is missing.
'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------

Private PUB_ERROR_VAL As Double
Private PUB_DF_VAL As Double
Private PUB_PB_VAL As Double

Private PUB_SPOT_VAL As Double
Private PUB_MODEL_VAL As Double

Private PUB_IBT_TREE_MAT As Variant
Private PUB_IBT_COEF_MAT As Variant

Private PUB_STRIKE_VEC As Variant
Private PUB_BID_MARKET_VEC As Variant
Private PUB_ASK_MARKET_VEC As Variant

Private PUB_BID_MODEL_VEC As Variant
Private PUB_ASK_MODEL_VEC As Variant
Private PUB_MID_MODEL_VEC As Variant

Private Const PUB_EPSILON As Double = 2 ^ 52


'************************************************************************************
'************************************************************************************
'FUNCTION      : IBT_VALUATION_FUNC

'DESCRIPTION   : An IBT is a generalization of the Cox, Ross, and Rubinstein
'binomial tree (CRR) for option pricing (CRR [1979]). IBT techniques, like the
'CRR technique, build a binomial tree to describe the evolution of the values
'of an underlying asset. An IBT differs from CRR because the probabilities
'attached to outcomes in the tree are inferred from a collection of actual
'option prices, rather than simply deduced from the behavior of the underlying
'asset. These optionimplied risk-neutral probabilities (or alternatively, the
'closely related risk-neutral state-contingent claim prices) are then available
'to be used to price other options. Stephen Ross asserts that options should be
'spanned by state-contingent claims (Ross [1976]). One implication is that with
'sufficient structure, we should be able to infer state-contingent claim prices
'or a probability density from options prices (Rubinstein [1994, p779]).

'LIBRARY       : DERIVATIVES
'GROUP         : IBT
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/31/2009
'REFERENCES    :
'http://www.mathfinance.cn/implied-binomial-tree/
'http://papers.ssrn.com/sol3/papers.cfm?abstract_id=541744
'http://home.wlu.edu/~schwartza/proofcrack.pdf
'http://www.quantonline.co.za/documents/Generating%20South%20African%20Volatility%20Surface.pdf
'************************************************************************************
'************************************************************************************

Function IBT_VALUATION_FUNC(ByVal S_VAL As Double, _
ByVal V_VAL As Double, _
ByVal T_VAL As Double, _
ByVal RF_VAL As Double, _
ByVal N_VAL As Long, _
ByRef STRIKE_RNG As Variant, _
ByRef ASK_RNG As Variant, _
ByRef BID_RNG As Variant, _
ByRef PARAM_RNG As Variant, _
Optional ByRef VOLAT_RNG As Variant, _
Optional ByVal OPTION_FLAG As Integer = 1, _
Optional ByVal MIN_IMP_VAL As Double = 0#, _
Optional ByVal MAX_IMP_VAL As Double = 1#, _
Optional ByVal CND_TYPE As Integer = 0, _
Optional ByVal OUTPUT As Integer = 0, _
Optional ByVal nLOOPS As Long = 1000, _
Optional ByVal tolerance As Double = 10 ^ -10, _
Optional ByVal epsilon As Double = 10 ^ -5)

'Option Strikes, Asks, and Bids (from a broker's web site at close of trade)

Dim i As Long
Dim j As Long
Dim k As Long
Dim NROWS As Long

Dim ASK_VECTOR As Variant
Dim BID_VECTOR As Variant

Dim NODES_ARR As Variant
Dim PARAM_VECTOR As Variant
Dim VOLAT_VECTOR As Variant
Dim STRIKE_VECTOR As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

STRIKE_VECTOR = STRIKE_RNG
If UBound(STRIKE_VECTOR, 1) = 1 Then: STRIKE_VECTOR = MATRIX_TRANSPOSE_FUNC(STRIKE_VECTOR)
NROWS = UBound(STRIKE_VECTOR, 1)

ASK_VECTOR = ASK_RNG
If UBound(ASK_VECTOR, 1) = 1 Then: ASK_VECTOR = MATRIX_TRANSPOSE_FUNC(ASK_VECTOR)
If NROWS <> UBound(ASK_VECTOR, 1) Then: GoTo ERROR_LABEL

BID_VECTOR = BID_RNG
If UBound(BID_VECTOR, 1) = 1 Then: BID_VECTOR = MATRIX_TRANSPOSE_FUNC(BID_VECTOR)
If NROWS <> UBound(BID_VECTOR, 1) Then: GoTo ERROR_LABEL

'------------------------------------------------------------------------------------
PUB_IBT_TREE_MAT = IBT_BINOMIAL_TREE_FUNC(S_VAL, V_VAL, T_VAL, RF_VAL, N_VAL)
PUB_SPOT_VAL = S_VAL
PUB_DF_VAL = Exp(-RF_VAL * T_VAL)

PUB_STRIKE_VEC = STRIKE_VECTOR
PUB_BID_MARKET_VEC = BID_VECTOR
PUB_ASK_MARKET_VEC = ASK_VECTOR

ReDim NODES_ARR(1 To NROWS * 2, 1 To (N_VAL + 1) + 2)

k = 1
For j = 1 To NROWS * 2 Step 2
    For i = 1 To N_VAL + 1
        NODES_ARR(j, i) = MAXIMUM_FUNC(PUB_IBT_TREE_MAT(i, N_VAL + 1) - PUB_STRIKE_VEC(k, 1), 0) * PUB_DF_VAL
        NODES_ARR(j + 1, i) = NODES_ARR(j, i)
    Next i
    k = k + 1
Next j

k = 1
For j = 1 To NROWS * 2 Step 2
    NODES_ARR(j, N_VAL + 2) = "<="
    NODES_ARR(j + 1, N_VAL + 2) = ">="
    
    NODES_ARR(j, N_VAL + 3) = ASK_VECTOR(k, 1)
    NODES_ARR(j + 1, N_VAL + 3) = BID_VECTOR(k, 1)
    k = k + 1
Next j

PUB_IBT_COEF_MAT = NODES_ARR
ReDim PUB_BID_MODEL_VEC(1 To NROWS, 1 To 1)
ReDim PUB_ASK_MODEL_VEC(1 To NROWS, 1 To 1)
ReDim PUB_MID_MODEL_VEC(1 To NROWS, 1 To 1)

'------------------------------------------------------------------------------------
If IsArray(PARAM_RNG) = True Then
'-----------------------------------------------------------------------------------
    PARAM_VECTOR = PARAM_RNG
    If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)
    Call IBT_OBJ1_FUNC(PARAM_VECTOR)
'-----------------------------------------------------------------------------------
Else 'Starting Value: CRR Ending Nodal Risk-Neutral Probabilities
'-----------------------------------------------------------------------------------
    ReDim PARAM_VECTOR(1 To N_VAL + 1, 1 To 1)
    For i = 1 To N_VAL + 1
        j = N_VAL - i + 1
        PARAM_VECTOR(i, 1) = COMBINATIONS_FUNC(N_VAL, j) * (PUB_PB_VAL ^ j) * ((1 - PUB_PB_VAL) ^ (N_VAL - j))
    Next i
    'PARAM_VECTOR = NELDER_MEAD_OPTIMIZATION_FRAME_FUNC("IBT_OBJ1_FUNC", PARAM_VECTOR, "", True, 0, nLOOPS, epsilon)
    PARAM_VECTOR = NELDER_MEAD_OPTIMIZATION3_FUNC("IBT_OBJ1_FUNC", PARAM_VECTOR, 1000, 10 ^ -10)
    If IsArray(PARAM_VECTOR) = False Then: GoTo ERROR_LABEL
'-----------------------------------------------------------------------------------
End If
'------------------------------------------------------------------------------------

'------------------------------------------------------------------------------------
Select Case OUTPUT
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
Case 0
'------------------------------------------------------------------------------------
    ReDim TEMP_MATRIX(1 To 6, 1 To NROWS + 1)
    TEMP_MATRIX(1, 1) = "IBT PRICING MODEL"
    TEMP_MATRIX(2, 1) = "STRIKE"
    TEMP_MATRIX(3, 1) = "BID MARKET"
    TEMP_MATRIX(4, 1) = "BID MODEL SPREAD"
    TEMP_MATRIX(5, 1) = "ASK MARKET"
    TEMP_MATRIX(6, 1) = "ASK MODEL SPREAD"
    For i = 1 To NROWS
        TEMP_MATRIX(1, i + 1) = PUB_MID_MODEL_VEC(i, 1)
        TEMP_MATRIX(2, i + 1) = PUB_STRIKE_VEC(i, 1)
        TEMP_MATRIX(3, i + 1) = PUB_BID_MARKET_VEC(i, 1)
        TEMP_MATRIX(4, i + 1) = PUB_BID_MODEL_VEC(i, 1)
        TEMP_MATRIX(5, i + 1) = PUB_ASK_MARKET_VEC(i, 1)
        TEMP_MATRIX(6, i + 1) = PUB_ASK_MODEL_VEC(i, 1)
    Next i
'------------------------------------------------------------------------------------
    IBT_VALUATION_FUNC = TEMP_MATRIX
'------------------------------------------------------------------------------------
Case 1
'------------------------------------------------------------------------------------
    '------------------------------------------------------------------------------------
    If IsArray(VOLAT_RNG) = True Then
    '------------------------------------------------------------------------------------
        VOLAT_VECTOR = VOLAT_RNG
        If UBound(VOLAT_VECTOR, 1) = 1 Then: VOLAT_VECTOR = MATRIX_TRANSPOSE_FUNC(VOLAT_VECTOR)
        For i = 1 To NROWS
            If VOLAT_VECTOR(i, 1) < tolerance Then: VOLAT_VECTOR(i, 1) = BLACK_SCHOLES_IMPLIED_VOLATILITY_FUNC((BID_VECTOR(i, 1) + ASK_VECTOR(i, 1)) / 2, S_VAL, STRIKE_VECTOR(i, 1), T_VAL, RF_VAL, RF_VAL - 0, OPTION_FLAG, MIN_IMP_VAL, MAX_IMP_VAL, CND_TYPE)
        Next i
    '------------------------------------------------------------------------------------
    Else
    '------------------------------------------------------------------------------------
1984:
        ReDim VOLAT_VECTOR(1 To NROWS, 1 To 1)
        For i = 1 To NROWS: VOLAT_VECTOR(i, 1) = BLACK_SCHOLES_IMPLIED_VOLATILITY_FUNC((BID_VECTOR(i, 1) + ASK_VECTOR(i, 1)) / 2, S_VAL, STRIKE_VECTOR(i, 1), T_VAL, RF_VAL, RF_VAL - 0, OPTION_FLAG, MIN_IMP_VAL, MAX_IMP_VAL, CND_TYPE): Next i
    '------------------------------------------------------------------------------------
    End If
    '------------------------------------------------------------------------------------
    ReDim TEMP_MATRIX(1 To 12, 1 To NROWS + 1)
    
    TEMP_MATRIX(1, 1) = "IBT PRICING MODEL"
    TEMP_MATRIX(2, 1) = "STRIKE"
    
    TEMP_MATRIX(3, 1) = "BID MARKET"
    TEMP_MATRIX(4, 1) = "BID MODEL SPREAD"
    
    TEMP_MATRIX(5, 1) = "ASK MARKET"
    TEMP_MATRIX(6, 1) = "ASK MODEL SPREAD"
    
    TEMP_MATRIX(7, 1) = "MID MARKET"
    TEMP_MATRIX(8, 1) = "MID MODEL SPREAD"
    
    TEMP_MATRIX(9, 1) = "IMPLIED V_VAL"
    TEMP_MATRIX(10, 1) = "BLACK THEORETICAL"
    TEMP_MATRIX(11, 1) = "TARGET MODEL"
    '% error of fit --> Error between mid-spread IBT option and Black-Scholes
    TEMP_MATRIX(12, 1) = "TARGET MID"
    '% error of fit --> Error between mid-spread ATM option and Black-Scholes
    
    For i = 1 To NROWS
        TEMP_MATRIX(1, i + 1) = PUB_MID_MODEL_VEC(i, 1)
        TEMP_MATRIX(2, i + 1) = PUB_STRIKE_VEC(i, 1)
        TEMP_MATRIX(3, i + 1) = PUB_BID_MARKET_VEC(i, 1)
        TEMP_MATRIX(4, i + 1) = PUB_BID_MODEL_VEC(i, 1)
        TEMP_MATRIX(5, i + 1) = PUB_ASK_MARKET_VEC(i, 1)
        TEMP_MATRIX(6, i + 1) = PUB_ASK_MODEL_VEC(i, 1)
        TEMP_MATRIX(7, i + 1) = (PUB_BID_MARKET_VEC(i, 1) + PUB_ASK_MARKET_VEC(i, 1)) / 2
        TEMP_MATRIX(8, i + 1) = (PUB_BID_MODEL_VEC(i, 1) + PUB_ASK_MODEL_VEC(i, 1)) / 2
        TEMP_MATRIX(9, i + 1) = VOLAT_VECTOR(i, 1)
        TEMP_MATRIX(10, i + 1) = GENERALIZED_BLACK_SCHOLES_FUNC(S_VAL, PUB_STRIKE_VEC(i, 1), T_VAL, RF_VAL, RF_VAL - 0, VOLAT_VECTOR(i, 1), OPTION_FLAG, CND_TYPE)
        If TEMP_MATRIX(1, i + 1) <> 0 Then
            TEMP_MATRIX(11, i + 1) = TEMP_MATRIX(10, i + 1) / TEMP_MATRIX(1, i + 1) - 1
        Else
            TEMP_MATRIX(11, i + 1) = CVErr(xlErrNA)
        End If
        If TEMP_MATRIX(7, i + 1) <> 0 Then
            TEMP_MATRIX(12, i + 1) = TEMP_MATRIX(10, i + 1) / TEMP_MATRIX(7, i + 1) - 1
        Else
            TEMP_MATRIX(12, i + 1) = CVErr(xlErrNA)
        End If
    Next i
'------------------------------------------------------------------------------------
    IBT_VALUATION_FUNC = TEMP_MATRIX
'------------------------------------------------------------------------------------'------------------------------------------------------------------------------------
Case 2
'------------------------------------------------------------------------------------
    IBT_VALUATION_FUNC = PUB_IBT_TREE_MAT
'------------------------------------------------------------------------------------
Case 3
'------------------------------------------------------------------------------------
    IBT_VALUATION_FUNC = PUB_IBT_COEF_MAT
'------------------------------------------------------------------------------------
Case 4
'------------------------------------------------------------------------------------
    IBT_VALUATION_FUNC = PUB_MID_MODEL_VEC
'------------------------------------------------------------------------------------
Case 5
'------------------------------------------------------------------------------------
    IBT_VALUATION_FUNC = PUB_BID_MODEL_VEC
'------------------------------------------------------------------------------------
Case 6
'------------------------------------------------------------------------------------
    IBT_VALUATION_FUNC = PUB_ASK_MODEL_VEC
'------------------------------------------------------------------------------------
Case 7
'------------------------------------------------------------------------------------
    IBT_VALUATION_FUNC = PUB_MODEL_VAL
'------------------------------------------------------------------------------------
Case 8
'------------------------------------------------------------------------------------
    IBT_VALUATION_FUNC = PUB_ERROR_VAL
'------------------------------------------------------------------------------------
Case 9
'------------------------------------------------------------------------------------
    IBT_VALUATION_FUNC = PARAM_VECTOR
'------------------------------------------------------------------------------------
Case Else
'------------------------------------------------------------------------------------
    IBT_VALUATION_FUNC = IBT_CONSTRAINT_FUNC(PARAM_VECTOR)
'------------------------------------------------------------------------------------
End Select
'------------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
IBT_VALUATION_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : IBT_BINOMIAL_TREE_FUNC
'DESCRIPTION   : Function to build a binomial tree to describe the evolution
'of the values of an underlying asset.
'LIBRARY       : DERIVATIVES
'GROUP         : IBT
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/31/2009
'************************************************************************************
'************************************************************************************

Private Function IBT_BINOMIAL_TREE_FUNC(ByVal S_VAL As Double, _
ByVal V_VAL As Double, _
ByVal T_VAL As Double, _
ByVal RF_VAL As Double, _
ByVal N_VAL As Long)

Dim i As Long
Dim j As Long

Dim DN_VAL As Double
Dim PB_VAL As Double
Dim DF_VAL As Double
Dim UP_VAL As Double

Dim DT_VAL As Double
Dim NODES_ARR As Variant

On Error GoTo ERROR_LABEL

DT_VAL = T_VAL / N_VAL
DN_VAL = Exp(-Sqr(DT_VAL) * V_VAL) 'Per Step
DF_VAL = Exp(DT_VAL * RF_VAL) 'Per Step
UP_VAL = Exp(Sqr(DT_VAL) * V_VAL) 'Per Step
PB_VAL = (DF_VAL - DN_VAL) / (UP_VAL - DN_VAL) 'Per Step
PUB_PB_VAL = PB_VAL

ReDim NODES_ARR(0 To N_VAL + 1, 0 To N_VAL + 1)
NODES_ARR(0, 0) = "IBT"
For i = 1 To N_VAL + 1
    If i = 1 Then NODES_ARR(i, i) = S_VAL Else NODES_ARR(i, i) = NODES_ARR(i - 1, i - 1) * DN_VAL
    For j = i + 1 To N_VAL + 1
        NODES_ARR(j, i) = ""
        NODES_ARR(i, j) = NODES_ARR(i, j - 1) * UP_VAL
    Next j
    NODES_ARR(0, i) = i - 1
    NODES_ARR(N_VAL + 1 - i + 1, 0) = NODES_ARR(0, i)
Next i

'-------------------------------------------------------------------------
IBT_BINOMIAL_TREE_FUNC = NODES_ARR
'-------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
IBT_BINOMIAL_TREE_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : IBT_OBJ1_FUNC
'DESCRIPTION   : Objective function for IBT Ending Nodal Risk-Neutral Probabilities
'LIBRARY       : DERIVATIVES
'GROUP         : IBT
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/31/2009
'************************************************************************************
'************************************************************************************

Private Function IBT_OBJ1_FUNC(ByRef PARAM_VECTOR As Variant)

Dim i As Long
Dim j As Long
Dim k As Long

Dim NSIZE As Long
Dim NROWS As Long

Dim N_VAL As Long

Dim TEMP1_SUM As Double
Dim TEMP2_SUM As Double
Dim TEMP_FACTOR As Double

On Error GoTo ERROR_LABEL

NSIZE = UBound(PARAM_VECTOR, 1) 'Imp Prob
NROWS = UBound(PUB_STRIKE_VEC, 1)
N_VAL = NSIZE - 1

k = 1
For j = 1 To NROWS
    TEMP1_SUM = 0: For i = 1 To NSIZE: TEMP1_SUM = TEMP1_SUM + (PUB_IBT_COEF_MAT(k, i) * PARAM_VECTOR(i, 1)): Next i
    PUB_MID_MODEL_VEC(j, 1) = TEMP1_SUM 'IBT Option Price
    PUB_ASK_MODEL_VEC(j, 1) = PUB_ASK_MARKET_VEC(j, 1) - PUB_MID_MODEL_VEC(j, 1)
    PUB_BID_MODEL_VEC(j, 1) = PUB_MID_MODEL_VEC(j, 1) - PUB_BID_MARKET_VEC(j, 1)
    k = k + 2
Next j

'-------------------------------------------------------------------------------
TEMP1_SUM = 0: TEMP2_SUM = 0
For i = 1 To NSIZE
    j = N_VAL - i + 1
    TEMP_FACTOR = COMBINATIONS_FUNC(N_VAL, j) * (PUB_PB_VAL ^ j) * ((1 - PUB_PB_VAL) ^ (N_VAL - j))
    TEMP1_SUM = TEMP1_SUM + (PARAM_VECTOR(i, 1) - TEMP_FACTOR) ^ 2
    TEMP2_SUM = TEMP2_SUM + (PUB_IBT_TREE_MAT(i, N_VAL + 1) * PARAM_VECTOR(i, 1))
Next i
'-------------------------------------------------------------------------------
PUB_MODEL_VAL = TEMP2_SUM * PUB_DF_VAL
'-------------------------------------------------------------------------------
If IBT_CONSTRAINT_FUNC(PARAM_VECTOR) = True Then
    TEMP2_SUM = 1
Else
    TEMP2_SUM = PUB_EPSILON
End If
IBT_OBJ1_FUNC = TEMP1_SUM * TEMP2_SUM
'-------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
PUB_ERROR_VAL = PUB_EPSILON
IBT_OBJ1_FUNC = PUB_EPSILON
End Function



'************************************************************************************
'************************************************************************************
'FUNCTION      : IBT_OBJ2_FUNC
'DESCRIPTION   : Jackwerth and Rubinstein 96 objective function
'LIBRARY       : DERIVATIVES
'GROUP         : IBT
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/31/2009
'************************************************************************************
'************************************************************************************

Private Function IBT_OBJ2_FUNC(ByRef PARAM_VECTOR As Variant)

Dim i As Long
Dim NSIZE As Long

Dim TEMP1_SUM As Double
Dim TEMP2_SUM As Double
Dim epsilon As Double

On Error GoTo ERROR_LABEL

epsilon = 0 '10 ^ -15
NSIZE = UBound(PARAM_VECTOR, 1) 'Imp Prob

TEMP1_SUM = ((PARAM_VECTOR(2, 1) - 2 * PARAM_VECTOR(1, 1) + epsilon) ^ 2)

For i = 2 To NSIZE - 1
    TEMP1_SUM = TEMP1_SUM + ((PARAM_VECTOR(i + 1, 1) - 2 * PARAM_VECTOR(i, 1) + PARAM_VECTOR(i - 1, 1)) ^ 2)
Next i

TEMP1_SUM = TEMP1_SUM + ((epsilon - 2 * PARAM_VECTOR(NSIZE, 1) + PARAM_VECTOR(NSIZE - 1, 1)) ^ 2)

If IBT_CONSTRAINT_FUNC(PARAM_VECTOR) = True Then
    TEMP2_SUM = 1
Else
    TEMP2_SUM = PUB_EPSILON
End If
IBT_OBJ2_FUNC = TEMP1_SUM * TEMP2_SUM

Exit Function
ERROR_LABEL:
IBT_OBJ2_FUNC = PUB_EPSILON
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : IBT_CONSTRAINT_FUNC
'DESCRIPTION   : Constraint function for IBT Ending Nodal Risk-Neutral Probabilities
'LIBRARY       : DERIVATIVES
'GROUP         : IBT
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/31/2009
'************************************************************************************
'************************************************************************************

Private Function IBT_CONSTRAINT_FUNC(ByRef PARAM_VECTOR As Variant)

Dim i As Long

Dim NSIZE As Long
Dim NROWS As Long

Dim TEMP_SUM As Double

Const tolerance As Double = 0.000001

On Error GoTo ERROR_LABEL

IBT_CONSTRAINT_FUNC = True
'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
NSIZE = UBound(PARAM_VECTOR, 1)
NROWS = UBound(PUB_STRIKE_VEC, 1)
'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
For i = 1 To NROWS
    If PUB_BID_MODEL_VEC(i, 1) < 0 Or PUB_ASK_MODEL_VEC(i, 1) < 0 Then
         IBT_CONSTRAINT_FUNC = False
         Exit Function
    End If
Next i
'-----------------------------------------------------------------------------------
TEMP_SUM = 0
For i = 1 To NSIZE
    If PARAM_VECTOR(i, 1) < tolerance Then
        IBT_CONSTRAINT_FUNC = False
        Exit Function
    End If
    TEMP_SUM = TEMP_SUM + PARAM_VECTOR(i, 1)
Next i
'-----------------------------------------------------------------------------------
If Abs(TEMP_SUM - 1) > 0 Then
    IBT_CONSTRAINT_FUNC = False
    Exit Function
End If
'-----------------------------------------------------------------------------------
If Abs(PUB_MODEL_VAL - PUB_SPOT_VAL) > 0 Then
    IBT_CONSTRAINT_FUNC = False
    Exit Function
End If
'-----------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
IBT_CONSTRAINT_FUNC = False
End Function


