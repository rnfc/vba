Attribute VB_Name = "FINAN_DERIV_BINOM_VOL_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : RNG_LOCAL_VOLATILITY_TREE_FUNC
'DESCRIPTION   : This routine creates and plots a local volatility tree. Based on
'paper The Volatility. Smile and Its Implied Tree by Emanuel Derman, Iraj Kani

'LIBRARY       : DERIVATIVES
'GROUP         : BINOMIAL
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function RNG_LOCAL_VOLATILITY_TREE_FUNC(ByVal SPOT As Double, _
ByVal RISK_FREE_RATE As Double, _
ByVal EXPIRATION As Double, _
ByVal VOLATILITY As Double, _
ByVal SKEW As Double, _
ByVal TIME_STEPS As Double, _
Optional ByVal EXERCISE_TYPE As String = "e", _
Optional ByRef DST_WBOOK As Excel.Workbook)

'RISK_FREE_RATE = annualized risk free rate
'EXPIRATION = no of years
'SPOT = current price
'VOLATILITY = annual volatility
'SKEW = linear increase in implied vol percentage point for decrease in strike
'TIME_STEPS = time steps
'EXERCISE_TYPE = "e" 'consider only european exercise

'Debug.Print RNG_LOCAL_VOLATILITY_TREE_FUNC(100, 0.029559, 5, 0.1, 0.0005, 5, "e")

Dim i As Long
Dim j As Long
Dim k As Long

Dim TEMP_SUM As Double

Dim S0_VAL As Double
Dim DT_VAL As Double

Dim K2_VAL As Double
Dim V2_VAL As Double
Dim T2_VAL As Double
Dim N2_VAL As Double
Dim CP_VAL As Double
Dim FI_VAL As Double
Dim PP_VAL As Double

Dim MID_VAL As Double
Dim LAST_VAL As Double

Dim NUMER_VAL As Double
Dim DENOM_VAL As Double

Dim CTEMP_ARR As Variant
Dim ATEMP_ARR As Variant
Dim BTEMP_ARR As Variant

Dim STEMP_MATRIX As Variant
Dim ATEMP_MATRIX As Variant
Dim PTEMP_MATRIX As Variant

Dim OPTION_TYPE As String
Dim DST_WSHEET As Excel.Worksheet

On Error GoTo ERROR_LABEL

If DST_WBOOK Is Nothing Then Set DST_WBOOK = ActiveWorkbook

S0_VAL = SPOT
DT_VAL = EXPIRATION / TIME_STEPS

ReDim STEMP_MATRIX(1 To TIME_STEPS + 1, 1 To TIME_STEPS + 1)
STEMP_MATRIX(1, 1) = S0_VAL
ReDim ATEMP_MATRIX(1 To TIME_STEPS + 1, 1 To TIME_STEPS + 1)
ATEMP_MATRIX(1, 1) = 1
ReDim PTEMP_MATRIX(1 To TIME_STEPS, 1 To TIME_STEPS)

For i = 1 To TIME_STEPS
'Step 1 : Find Center nodes of stock tree (EQ 5)
    If ((i + 1) Mod 2) = 0 Then
        'this is even level branch
        'let us find the upper node first (EQN 8)
        MID_VAL = Round((i + 1) / 2)
        SPOT = STEMP_MATRIX(MID_VAL, i)
        OPTION_TYPE = "c"
        K2_VAL = SPOT
        V2_VAL = VOLATILITY - (K2_VAL - S0_VAL) * SKEW
        T2_VAL = i * DT_VAL
        'V2_VAL = VOLATILITY - (K2_VAL - S0_VAL * Exp(RISK_FREE_RATE * T2_VAL)) * SKEW
        N2_VAL = i
        CP_VAL = BINOMIAL_TREE_OPTION_PRICE_FUNC(S0_VAL, K2_VAL, T2_VAL, RISK_FREE_RATE, V2_VAL, N2_VAL, OPTION_TYPE, EXERCISE_TYPE)
        TEMP_SUM = 0
        For k = 1 To MID_VAL - 1
            TEMP_SUM = TEMP_SUM + ATEMP_MATRIX(k, i) * (Exp(RISK_FREE_RATE * DT_VAL) * STEMP_MATRIX(k, i) - STEMP_MATRIX(MID_VAL, i))
        Next k
        NUMER_VAL = SPOT * (Exp(RISK_FREE_RATE * DT_VAL) * CP_VAL + ATEMP_MATRIX(MID_VAL, i) * SPOT - TEMP_SUM)
        FI_VAL = SPOT * Exp(RISK_FREE_RATE * DT_VAL)
        DENOM_VAL = ATEMP_MATRIX(MID_VAL, i) * FI_VAL - Exp(RISK_FREE_RATE * DT_VAL) * CP_VAL + TEMP_SUM
        STEMP_MATRIX(MID_VAL, i + 1) = NUMER_VAL / DENOM_VAL
        'use formula in page 8 to find lower node
        STEMP_MATRIX(MID_VAL + 1, i + 1) = SPOT * SPOT / STEMP_MATRIX(MID_VAL, i + 1)
        LAST_VAL = MID_VAL + 1
    Else
        'this is odd level branch
        'simply set center point price to spot
        'tesmi = Application.WorksheetFunction.Ceiling(2.2, 1)
        MID_VAL = Application.WorksheetFunction.Ceiling((i + 1) / 2, 1)
        STEMP_MATRIX(MID_VAL, i + 1) = S0_VAL
        LAST_VAL = MID_VAL
    End If

    'Step 2 : Find upper nodes of stock tree (EQ 6)
    For j = MID_VAL - 1 To 1 Step -1
        OPTION_TYPE = "c"
        K2_VAL = STEMP_MATRIX(j, i)
        V2_VAL = VOLATILITY - (K2_VAL - S0_VAL) * SKEW
        T2_VAL = i * DT_VAL
        'V2_VAL = VOLATILITY - (K2_VAL - S0_VAL * Exp(RISK_FREE_RATE * T2_VAL)) * SKEW
        N2_VAL = i
        CP_VAL = BINOMIAL_TREE_OPTION_PRICE_FUNC(S0_VAL, K2_VAL, T2_VAL, RISK_FREE_RATE, V2_VAL, N2_VAL, OPTION_TYPE, EXERCISE_TYPE)
        TEMP_SUM = 0
        For k = 1 To j - 1
            TEMP_SUM = TEMP_SUM + ATEMP_MATRIX(k, i) * (Exp(RISK_FREE_RATE * DT_VAL) * STEMP_MATRIX(k, i) - STEMP_MATRIX(j, i))
        Next k
        FI_VAL = STEMP_MATRIX(j, i) * Exp(RISK_FREE_RATE * DT_VAL)
        NUMER_VAL = STEMP_MATRIX(j + 1, i + 1) * (Exp(RISK_FREE_RATE * DT_VAL) * CP_VAL - TEMP_SUM) - ATEMP_MATRIX(j, i) * STEMP_MATRIX(j, i) * (FI_VAL - STEMP_MATRIX(j + 1, i + 1))
        DENOM_VAL = Exp(RISK_FREE_RATE * DT_VAL) * CP_VAL - TEMP_SUM - ATEMP_MATRIX(j, i) * (FI_VAL - STEMP_MATRIX(j + 1, i + 1))
        STEMP_MATRIX(j, i + 1) = NUMER_VAL / DENOM_VAL
    Next j
    
    'Step 3 : Find lower nodes of stock tree (EQ 9)
    For j = LAST_VAL + 1 To (i + 1)
        OPTION_TYPE = "p"
        K2_VAL = STEMP_MATRIX(j - 1, i)
        V2_VAL = VOLATILITY - (K2_VAL - S0_VAL) * SKEW
        T2_VAL = i * DT_VAL
        'V2_VAL = VOLATILITY - (K2_VAL - S0_VAL * Exp(RISK_FREE_RATE * T2_VAL)) * SKEW
        N2_VAL = i
        PP_VAL = BINOMIAL_TREE_OPTION_PRICE_FUNC(S0_VAL, K2_VAL, T2_VAL, RISK_FREE_RATE, V2_VAL, N2_VAL, OPTION_TYPE, EXERCISE_TYPE)
        TEMP_SUM = 0
        For k = j + 1 To i + 1
            TEMP_SUM = TEMP_SUM + ATEMP_MATRIX(k - 1, i) * (STEMP_MATRIX(j - 1, i) - Exp(RISK_FREE_RATE * DT_VAL) * STEMP_MATRIX(k - 1, i))
        Next k
        FI_VAL = STEMP_MATRIX(j - 1, i) * Exp(RISK_FREE_RATE * DT_VAL)
        
        NUMER_VAL = STEMP_MATRIX(j - 1, i + 1) * (Exp(RISK_FREE_RATE * DT_VAL) * PP_VAL - TEMP_SUM) + ATEMP_MATRIX(j - 1, i) * STEMP_MATRIX(j - 1, i) * (FI_VAL - STEMP_MATRIX(j - 1, i + 1))
        DENOM_VAL = Exp(RISK_FREE_RATE * DT_VAL) * PP_VAL - TEMP_SUM + ATEMP_MATRIX(j - 1, i) * (FI_VAL - STEMP_MATRIX(j - 1, i + 1))
        STEMP_MATRIX(j, i + 1) = NUMER_VAL / DENOM_VAL
    Next j
    
    'Step 4 : Find nodes of probability and Arrow Debreu tree
    For j = 1 To i
        PTEMP_MATRIX(j, i) = (Exp(RISK_FREE_RATE * DT_VAL) * STEMP_MATRIX(j, i) - STEMP_MATRIX(j + 1, i + 1)) / (STEMP_MATRIX(j, i + 1) - STEMP_MATRIX(j + 1, i + 1))
        ATEMP_MATRIX(1, i + 1) = ATEMP_MATRIX(1, i) * PTEMP_MATRIX(1, i) / Exp(RISK_FREE_RATE * DT_VAL)
        ATEMP_MATRIX(i + 1, i + 1) = ATEMP_MATRIX(i, i) * (1 - PTEMP_MATRIX(i, i)) / Exp(RISK_FREE_RATE * DT_VAL)
        For k = 2 To i
            ATEMP_MATRIX(k, i + 1) = (ATEMP_MATRIX(k - 1, i) * (1 - PTEMP_MATRIX(k - 1, i)) + ATEMP_MATRIX(k, i) * PTEMP_MATRIX(k, i)) / Exp(RISK_FREE_RATE * DT_VAL)
        Next k
    Next j
Next i

ReDim LTEMP_MATRIX(1 To TIME_STEPS, 1 To TIME_STEPS)
For i = 1 To TIME_STEPS
    For j = 1 To i
        LTEMP_MATRIX(j, i) = (1 / DT_VAL ^ 0.5) * ((PTEMP_MATRIX(j, i) * (1 - PTEMP_MATRIX(j, i))) ^ 0.5 * Log(STEMP_MATRIX(j, i + 1) / STEMP_MATRIX(j + 1, i + 1)))
    Next j
Next i

ReDim CTEMP_ARR(1 To TIME_STEPS + 1)
For i = 1 To TIME_STEPS + 1
    ReDim ATEMP_ARR(1 To i)
    For j = 1 To i
        ReDim BTEMP_ARR(1 To 4)
        BTEMP_ARR(1) = STEMP_MATRIX(j, i)
        BTEMP_ARR(2) = ATEMP_MATRIX(j, i)
        If i < TIME_STEPS + 1 Then
            BTEMP_ARR(3) = PTEMP_MATRIX(j, i)
            BTEMP_ARR(4) = LTEMP_MATRIX(j, i)
        End If
        ATEMP_ARR(j) = BTEMP_ARR
    Next j
    CTEMP_ARR(i) = ATEMP_ARR
Next i

Set DST_WSHEET = WSHEET_ADD_FUNC(PARSE_CURRENT_TIME_FUNC("_"), DST_WBOOK)
RNG_LOCAL_VOLATILITY_TREE_FUNC = PRINT_MULTI_BINOMIAL_TREE_FUNC(CTEMP_ARR, DST_WSHEET)

Exit Function
ERROR_LABEL:
RNG_LOCAL_VOLATILITY_TREE_FUNC = False
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : PRINT_MULTI_BINOMIAL_TREE_FUNC

'DESCRIPTION   : Displays binomial tree nodes with multiple value on a worksheet
'DATA_GROUP - 3 level array

'LIBRARY       : DERIVATIVES
'GROUP         : BINOMIAL
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Private Function PRINT_MULTI_BINOMIAL_TREE_FUNC(ByRef DATA_GROUP As Variant, _
ByRef DST_WSHEET As Excel.Worksheet)

Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim ii As Long
Dim jj As Long
Dim kk As Long
Dim ll As Long

Dim NROWS As Long

Dim TEMP_ARR As Variant
Dim VALUES_ARR As Variant

Dim START_X_VAL As Double
Dim START_Y_VAL As Double

Dim END_X_VAL As Double
Dim END_Y_VAL As Double

Dim TEMP_RNG As Excel.Range
Dim SRC_SHAPE As Excel.Shape
'Dim DST_WSHEET As Excel.Worksheet

Dim END_NODE_TOP_CELL As Excel.Range
Dim END_NODE_BOTTOM_CELL As Excel.Range

Dim END_NODE_2TOP_CELL As Excel.Range
Dim END_NODE_2BOTTOM_CELL As Excel.Range

Dim START_NODE_TOP_CELL As Excel.Range
Dim START_NODE_BOTTOM_CELL As Excel.Range

On Error GoTo ERROR_LABEL

PRINT_MULTI_BINOMIAL_TREE_FUNC = False

Set TEMP_RNG = DST_WSHEET.Cells
TEMP_RNG.value = ""
TEMP_RNG.Borders(xlDiagonalDown).LineStyle = xlNone
TEMP_RNG.Borders(xlDiagonalUp).LineStyle = xlNone
TEMP_RNG.Borders(xlEdgeLeft).LineStyle = xlNone
TEMP_RNG.Borders(xlEdgeTop).LineStyle = xlNone
TEMP_RNG.Borders(xlEdgeBottom).LineStyle = xlNone
TEMP_RNG.Borders(xlEdgeRight).LineStyle = xlNone
TEMP_RNG.Borders(xlInsideVertical).LineStyle = xlNone
TEMP_RNG.Borders(xlInsideHorizontal).LineStyle = xlNone

For Each SRC_SHAPE In DST_WSHEET.Shapes
    SRC_SHAPE.Delete
Next SRC_SHAPE

With TEMP_RNG.Interior
    .ColorIndex = 2
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
End With

'determine the longest array of node values
l = UBound(DATA_GROUP, 1)
ll = 0
For i = 1 To l 'loop over time nodes
    TEMP_ARR = DATA_GROUP(i) 'array of node value for a given time
    If UBound(TEMP_ARR, 1) > ll Then
        ll = UBound(TEMP_ARR, 1)
    End If
Next i
'find NROWS
NROWS = (ll - 1) * 3 + 2
'need to find central node

jj = 1

TEMP_ARR = DATA_GROUP(1)
VALUES_ARR = TEMP_ARR(1)
kk = UBound(VALUES_ARR, 1)

For i = 1 To l 'loop over time nodes
    TEMP_ARR = DATA_GROUP(i) 'array of node value for a given time
    ii = (l - i) * (kk + 1) + 1
    
    For j = 1 To i
        VALUES_ARR = TEMP_ARR(j)
        For k = 1 To kk 'loop over values for a given node
            Set TEMP_RNG = DST_WSHEET.Cells(ii + k - 1, jj)
            TEMP_RNG.value = VALUES_ARR(k)
            TEMP_RNG.Borders(xlDiagonalDown).LineStyle = xlNone
            TEMP_RNG.Borders(xlDiagonalUp).LineStyle = xlNone
            With TEMP_RNG.Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .WEIGHT = xlThin
                .ColorIndex = xlAutomatic
            End With
            With TEMP_RNG.Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .WEIGHT = xlThin
                .ColorIndex = xlAutomatic
            End With
            With TEMP_RNG.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .WEIGHT = xlThin
                .ColorIndex = xlAutomatic
            End With
            With TEMP_RNG.Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .WEIGHT = xlThin
                .ColorIndex = xlAutomatic
            End With
        Next k
        
        If i < l Then
            Set START_NODE_TOP_CELL = DST_WSHEET.Cells(ii, jj)
            Set START_NODE_BOTTOM_CELL = DST_WSHEET.Cells(ii + kk - 1, jj)
            Set END_NODE_TOP_CELL = DST_WSHEET.Cells(ii - kk - 1, jj + 2)
            Set END_NODE_BOTTOM_CELL = DST_WSHEET.Cells(ii - 2, jj + 2)
            Set END_NODE_2TOP_CELL = DST_WSHEET.Cells(ii + kk + 1, jj + 2)
            Set END_NODE_2BOTTOM_CELL = DST_WSHEET.Cells(ii + 2 * kk, jj + 2)
            
            START_X_VAL = START_NODE_TOP_CELL.Left + START_NODE_TOP_CELL.Width
            START_Y_VAL = (START_NODE_TOP_CELL.Top + (START_NODE_BOTTOM_CELL.Top + START_NODE_BOTTOM_CELL.Height)) / 2

            END_X_VAL = END_NODE_TOP_CELL.Left
            END_Y_VAL = (END_NODE_TOP_CELL.Top + (END_NODE_BOTTOM_CELL.Top + END_NODE_BOTTOM_CELL.Height)) / 2

            Set SRC_SHAPE = DST_WSHEET.Shapes.AddLine(START_X_VAL, START_Y_VAL, END_X_VAL, END_Y_VAL)
            SRC_SHAPE.Line.EndArrowheadStyle = msoArrowheadTriangle
            SRC_SHAPE.Line.EndArrowheadLength = msoArrowheadLengthMedium
            SRC_SHAPE.Line.EndArrowheadWidth = msoArrowheadWidthMedium
            
            END_X_VAL = END_NODE_2TOP_CELL.Left
            END_Y_VAL = (END_NODE_2TOP_CELL.Top + (END_NODE_2BOTTOM_CELL.Top + END_NODE_2BOTTOM_CELL.Height)) / 2
            
            Set SRC_SHAPE = DST_WSHEET.Shapes.AddLine(START_X_VAL, START_Y_VAL, END_X_VAL, END_Y_VAL)
            SRC_SHAPE.Line.EndArrowheadStyle = msoArrowheadTriangle
            SRC_SHAPE.Line.EndArrowheadLength = msoArrowheadLengthMedium
            SRC_SHAPE.Line.EndArrowheadWidth = msoArrowheadWidthMedium
        End If
        ii = ii + kk + 3 + (kk - 1)
    Next j
    
    jj = jj + 2
Next i

PRINT_MULTI_BINOMIAL_TREE_FUNC = True

Exit Function
ERROR_LABEL:
PRINT_MULTI_BINOMIAL_TREE_FUNC = False
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : BINOMIAL_TREE_OPTION_PRICE_FUNC

'DESCRIPTION   : Calculates American/ European Call/Put prices using CRR Binomial tree.
'LIBRARY       : DERIVATIVES
'GROUP         : BINOMIAL
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Private Function BINOMIAL_TREE_OPTION_PRICE_FUNC(ByVal S_VAL As Double, _
ByVal K_VAL As Double, _
ByVal T_VAL As Double, _
ByVal RF_VAL As Double, _
ByVal V_VAL As Double, _
ByVal N_VAL As Double, _
ByVal OT_VAL As String, _
ByVal ET_VAL As String)

'OT_VAL - use "c" for call and "p" for put option
'S_VAL - S_VAL price
'K_VAL - option K_VAL
'T_VAL - option maturity
'RF_VAL - risk free RF_VAL
'V_VAL - V_VAL
'ET_VAL - use "a" for American and "e" for european
'N_VAL - no of time steps for the binomial tree

Dim i As Long
Dim j As Long
Dim k As Long
Dim S0_VAL As Double
Dim DT_VAL As Double

Dim UP_VAL As Double
Dim DN_VAL As Double

Dim P1_VAL As Double
Dim P2_VAL As Double

Dim PV_VAL As Double
Dim TM_VAL As Double

Dim TEMP1_MATRIX As Variant
Dim TEMP2_MATRIX As Variant

On Error GoTo ERROR_LABEL

S0_VAL = S_VAL
If OT_VAL = "c" Then k = 1 Else k = -1
DT_VAL = T_VAL / N_VAL
UP_VAL = Exp(V_VAL * DT_VAL ^ 0.5) 'size of up jump
DN_VAL = Exp(-V_VAL * DT_VAL ^ 0.5) 'size of down jump
P1_VAL = (UP_VAL - Exp(RF_VAL * DT_VAL)) / (UP_VAL - DN_VAL) 'probability of up jump
P2_VAL = 1 - P1_VAL 'probability of down jump

ReDim TEMP1_MATRIX(1 To N_VAL + 1, 1 To N_VAL + 1) 'hlods stock prices
TEMP1_MATRIX(1, 1) = S0_VAL
For i = 1 To UBound(TEMP1_MATRIX, 1) - 1
    TEMP1_MATRIX(1, i + 1) = TEMP1_MATRIX(1, i) * Exp(V_VAL * DT_VAL ^ 0.5)
    For j = 2 To i + 1
        TEMP1_MATRIX(j, i + 1) = TEMP1_MATRIX(j - 1, i) * Exp(-V_VAL * DT_VAL ^ 0.5)
    Next j
Next i

ReDim TEMP2_MATRIX(1 To N_VAL + 1, 1 To N_VAL + 1)
For i = 1 To N_VAL + 1
    TEMP2_MATRIX(i, N_VAL + 1) = MAXIMUM_FUNC(k * (TEMP1_MATRIX(i, N_VAL + 1) - K_VAL), 0)
Next i

For i = UBound(TEMP1_MATRIX, 2) - 1 To 1 Step -1
    For j = 1 To i
        PV_VAL = Exp(-RF_VAL * DT_VAL) * (P2_VAL * TEMP2_MATRIX(j, i + 1) + P1_VAL * TEMP2_MATRIX(j + 1, i + 1))
        TM_VAL = k * (TEMP1_MATRIX(j, i) - K_VAL)
        If ET_VAL = "a" Then
            TEMP2_MATRIX(j, i) = MAXIMUM_FUNC(PV_VAL, TM_VAL)
        Else
            TEMP2_MATRIX(j, i) = MAXIMUM_FUNC(PV_VAL, 0)
        End If
    Next j
Next i
BINOMIAL_TREE_OPTION_PRICE_FUNC = TEMP2_MATRIX(1, 1)

Exit Function
ERROR_LABEL:
BINOMIAL_TREE_OPTION_PRICE_FUNC = Err.number
End Function


