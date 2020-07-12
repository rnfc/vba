Attribute VB_Name = "FINAN_FI_BOND_DURATION_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
                    

'************************************************************************************
'************************************************************************************
'FUNCTION      : CONVEXITY_DURATION_FUNC

'DESCRIPTION   : Returns the Macauley duration and Convexity Table of a bond.
'Duration and Convexity are just the weighted average of the present
'value of the cash flows and is used as a measure of a bond price's
'response to changes in yield.

'LIBRARY       : BOND
'GROUP         : DURATION
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************

Function BOND_CONVEXITY_DURATION_FUNC(ByVal SETTLEMENT As Date, _
ByVal MATURITY As Date, _
ByVal COUPON As Double, _
ByVal YIELD As Double, _
Optional ByVal FREQUENCY As Integer = 2, _
Optional ByVal REDEMPTION As Double = 100, _
Optional ByVal COUNT_BASIS As Integer = 0, _
Optional ByVal OUTPUT As Integer = 0)

Dim i As Long
Dim j As Long
Dim k As Long

Dim Y_VAL As Double
Dim T_VAL As Double
Dim PM_VAL As Double
Dim DF_VAL As Double
Dim PV_VAL As Double

Dim PVW_VAL As Double
Dim DUR_VAL As Double
Dim CON_VAL As Double

Dim TEMP1_SUM As Double
Dim TEMP2_SUM As Double
Dim TEMP3_SUM As Double

Dim TENOR_VECTOR As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

If (MATURITY < SETTLEMENT) Then
    BOND_CONVEXITY_DURATION_FUNC = 0
    Exit Function
End If

k = FREQUENCY
If k = 0 Then: k = 1

Y_VAL = YIELD
TENOR_VECTOR = BOND_DATES_BOND_TENOR_FUNC(SETTLEMENT, MATURITY, k, COUNT_BASIS)
j = UBound(TENOR_VECTOR, 1)

'-----------------------------------------------------------------------------------
Select Case OUTPUT
'-----------------------------------------------------------------------------------
Case 0 'CONVEXITY / MODIFIED DURATION / DURATION / BOND CASH PRICE
'-----------------------------------------------------------------------------------
    i = j: GoSub PV_LINE
    '-------------------------first pass to calculate PV of coupons-----------------
    TEMP1_SUM = PM_VAL * DF_VAL
    For i = j - 1 To 1 Step -1
        GoSub PV_LINE
        TEMP1_SUM = PV_VAL + TEMP1_SUM
    Next i
    '---------------second pass to calculate duration and convexity -----------------
    TEMP2_SUM = 0: TEMP3_SUM = 0
    For i = j To 1 Step -1
        GoSub DUR_LINE
        TEMP2_SUM = TEMP2_SUM + DUR_VAL 'duration
        TEMP3_SUM = TEMP3_SUM + CON_VAL
    Next i
    'Convexity/MDuration/Duration/Bond Price
    BOND_CONVEXITY_DURATION_FUNC = Array(TEMP3_SUM, TEMP2_SUM / (1 + (Y_VAL / k)), TEMP2_SUM, TEMP1_SUM)
'-----------------------------------------------------------------------------------
Case Else
'-----------------------------------------------------------------------------------
    ReDim TEMP_MATRIX(0 To j, 1 To 7)
    TEMP_MATRIX(0, 1) = "TENOR"
    TEMP_MATRIX(0, 2) = "PAYMENTS"
    TEMP_MATRIX(0, 3) = "DISCOUNT FACTORS"
    TEMP_MATRIX(0, 4) = "PV PAYMENTS"
    TEMP_MATRIX(0, 5) = "PV WEIGHTS"
    TEMP_MATRIX(0, 6) = "DURATION"
    TEMP_MATRIX(0, 7) = "CONVEXITY"
    
    i = j: GoSub PV_LINE
    TEMP_MATRIX(i, 1) = T_VAL: TEMP_MATRIX(i, 2) = PM_VAL
    TEMP_MATRIX(i, 3) = DF_VAL: TEMP_MATRIX(i, 4) = PM_VAL * DF_VAL
    '-------------------------first pass to calculate PV of coupons-----------------
    TEMP1_SUM = PM_VAL * DF_VAL
    For i = j - 1 To 1 Step -1
        GoSub PV_LINE
        TEMP1_SUM = PV_VAL + TEMP1_SUM
        TEMP_MATRIX(i, 1) = T_VAL
        TEMP_MATRIX(i, 2) = PM_VAL
        TEMP_MATRIX(i, 3) = DF_VAL
        TEMP_MATRIX(i, 4) = PV_VAL
    Next i
    '---------------second pass to calculate duration and convexity -----------------
    For i = j To 1 Step -1
        GoSub DUR_LINE
        TEMP_MATRIX(i, 5) = PVW_VAL
        TEMP_MATRIX(i, 6) = DUR_VAL
        TEMP_MATRIX(i, 7) = CON_VAL 'convexity
    Next i
    BOND_CONVEXITY_DURATION_FUNC = TEMP_MATRIX
'-----------------------------------------------------------------------------------
End Select
'-----------------------------------------------------------------------------------

'epsilon = YIELD * 0.01
'P1_VAL = BOND_CASH_PRICE_FUNC(SETTLEMENT, MATURITY, COUPON, YIELD - epsilon, FREQUENCY, REDEMPTION, COUNT_BASIS)
'P2_VAL = BOND_CASH_PRICE_FUNC(SETTLEMENT, MATURITY, COUPON, YIELD, FREQUENCY, REDEMPTION, COUNT_BASIS)
'P3_VAL = BOND_CASH_PRICE_FUNC(SETTLEMENT, MATURITY, COUPON, YIELD + epsilon, FREQUENCY, REDEMPTION, COUNT_BASIS)
'P4_VAL = (P3_VAL + P1_VAL - 2 * P2_VAL) / (epsilon * epsilon)

'BOND_CONVEXITY_FUNC = P4_VAL / P2_VAL

'P1_VAL = BOND_CASH_PRICE_FUNC(SETTLEMENT, MATURITY, COUPON, YIELD - epsilon, FREQUENCY, REDEMPTION, COUNT_BASIS)
'P2_VAL = BOND_CASH_PRICE_FUNC(SETTLEMENT, MATURITY, COUPON, YIELD, FREQUENCY, COUNT_BASIS, REDEMPTION)
'P3_VAL = BOND_CASH_PRICE_FUNC(SETTLEMENT, MATURITY, COUPON, YIELD + epsilon, FREQUENCY, REDEMPTION, COUNT_BASIS)
'P4_VAL = (P3_VAL - P1_VAL) / (2 * epsilon)

'BOND_DURATION_FUNC = -P4_VAL / P2_VAL

Exit Function
'-----------------------------------------------------------------------------------------
PV_LINE:
'-----------------------------------------------------------------------------------------
    T_VAL = TENOR_VECTOR(i, 1)
    PM_VAL = (COUPON / k) * 100 + IIf(i = j, REDEMPTION, 0) 'PAYMENT
    DF_VAL = 1 / (1 + Y_VAL / k) ^ (k * T_VAL) 'DISC_FACTOR
    PV_VAL = PM_VAL * DF_VAL 'PV OF PAYMENTS
'-----------------------------------------------------------------------------------------
Return
'-----------------------------------------------------------------------------------------
DUR_LINE:
'-----------------------------------------------------------------------------------------
    GoSub PV_LINE
    PVW_VAL = PV_VAL / TEMP1_SUM
    DUR_VAL = PVW_VAL * T_VAL
    CON_VAL = T_VAL * ((1 / k) + T_VAL) * PM_VAL * DF_VAL / (TEMP1_SUM * (1 + Y_VAL / (k)) ^ k)
'-----------------------------------------------------------------------------------------
Return
'-----------------------------------------------------------------------------------------
ERROR_LABEL:
BOND_CONVEXITY_DURATION_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : YIELD_SENSITIVITY_FUNC
'DESCRIPTION   : YIELD SENSITIVITY TABLE
'LIBRARY       : BOND
'GROUP         : DURATION
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************

Function YIELD_SENSITIVITY_FUNC(ByVal OUTPUT As Integer, _
ByVal CLEAN_PRICE As Double, _
ByVal SETTLEMENT As Date, _
ByVal MATURITY As Date, _
ByVal COUPON As Double, _
ByVal MIN_YIELD As Double, _
ByVal MAX_YIELD As Double, _
ByVal DELTA_YIELD As Double, _
Optional ByVal FREQUENCY As Integer = 2, _
Optional ByVal REDEMPTION As Double = 100, _
Optional ByVal COUNT_BASIS As Integer = 0, _
Optional ByVal GUESS_YIELD As Double = 0.1)

Dim i As Long
Dim NROWS As Long

Dim PYIELD_VAL As Double
Dim CONVEXITY_VAL As Double
Dim MDURATION_VAL As Double

Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

If MIN_YIELD >= MAX_YIELD Then: GoTo ERROR_LABEL
If DELTA_YIELD = 0 Then: GoTo ERROR_LABEL

PYIELD_VAL = BOND_YIELD_FUNC(CLEAN_PRICE, SETTLEMENT, MATURITY, COUPON, FREQUENCY, REDEMPTION, COUNT_BASIS, GUESS_YIELD)
TEMP_MATRIX = BOND_CONVEXITY_DURATION_FUNC(SETTLEMENT, MATURITY, COUPON, PYIELD_VAL, FREQUENCY, REDEMPTION, COUNT_BASIS, 0)
CONVEXITY_VAL = TEMP_MATRIX(LBound(TEMP_MATRIX) + 0)
MDURATION_VAL = TEMP_MATRIX(LBound(TEMP_MATRIX) + 1)

'-----------------------------------------------------------------------------------------------------------------------------
If OUTPUT <> 0 Then
'-----------------------------------------------------------------------------------------------------------------------------
    ReDim TEMP_MATRIX(1 To 3, 1 To 2)
    TEMP_MATRIX(1, 1) = "PERCENT YIELD"
    TEMP_MATRIX(2, 1) = "MODIFIED DURATION"
    TEMP_MATRIX(3, 1) = "CONVEXITY"
    
    TEMP_MATRIX(1, 2) = PYIELD_VAL
    TEMP_MATRIX(2, 2) = MDURATION_VAL
    TEMP_MATRIX(3, 2) = CONVEXITY_VAL
'-----------------------------------------------------------------------------------------------------------------------------
Else
'-----------------------------------------------------------------------------------------------------------------------------
    NROWS = ((MAX_YIELD - MIN_YIELD) / DELTA_YIELD)
    ReDim TEMP_MATRIX(0 To NROWS + 1, 1 To 9)
    TEMP_MATRIX(0, 1) = "NEW YIELD"
    TEMP_MATRIX(0, 2) = "ESTIMATED CLEAN PRICE"
    TEMP_MATRIX(0, 3) = "DURATION PREDICTED PRICE CHANGE"
    TEMP_MATRIX(0, 4) = "DURATION PREDICTED PRICE"
    TEMP_MATRIX(0, 5) = "ERROR"
    TEMP_MATRIX(0, 6) = "CONVEXITY ADJUSTMENT"
    TEMP_MATRIX(0, 7) = "CONVEXITY ADJUSTED PREDICTED PRICE CHANGE"
    TEMP_MATRIX(0, 8) = "CONVEXITY ADJ. PREDICTED PRICE"
    TEMP_MATRIX(0, 9) = "ERROR"
    TEMP_MATRIX(1, 1) = MIN_YIELD
    For i = 1 To NROWS + 1
        If i > 1 Then
            TEMP_MATRIX(i, 1) = TEMP_MATRIX(i - 1, 1) + DELTA_YIELD 'New Yield
        Else
            TEMP_MATRIX(i, 1) = MIN_YIELD 'New Yield
        End If
        TEMP_MATRIX(i, 2) = BOND_CASH_PRICE_FUNC(SETTLEMENT, MATURITY, COUPON, TEMP_MATRIX(i, 1), FREQUENCY, REDEMPTION, COUNT_BASIS, 1) 'Bond Clean Price
        TEMP_MATRIX(i, 3) = DELTA_DURATION_PRICE_FUNC(MDURATION_VAL, TEMP_MATRIX(i, 1), PYIELD_VAL) 'Duration_Predicted_Price_Change
        TEMP_MATRIX(i, 4) = CLEAN_PRICE * (1 + TEMP_MATRIX(i, 3)) 'Duration Predicted Price
        TEMP_MATRIX(i, 5) = (TEMP_MATRIX(i, 4) - TEMP_MATRIX(i, 2)) / TEMP_MATRIX(i, 2) 'Error
        TEMP_MATRIX(i, 6) = 0.5 * CONVEXITY_VAL * ((TEMP_MATRIX(i, 1) - PYIELD_VAL) ^ 2) 'Convexity Adjustment
        TEMP_MATRIX(i, 7) = DELTA_CONVEXITY_PRICE_FUNC(CONVEXITY_VAL, MDURATION_VAL, TEMP_MATRIX(i, 1), PYIELD_VAL) 'Convexity Adjusted Predicted Price Change
        TEMP_MATRIX(i, 8) = CLEAN_PRICE * (1 + TEMP_MATRIX(i, 7)) 'Convexity Adj. --> Predicted Price
        TEMP_MATRIX(i, 9) = (TEMP_MATRIX(i, 8) - TEMP_MATRIX(i, 2)) / TEMP_MATRIX(i, 2) 'Error
    Next i
'-----------------------------------------------------------------------------------------------------------------------------
End If
'-----------------------------------------------------------------------------------------------------------------------------
YIELD_SENSITIVITY_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
YIELD_SENSITIVITY_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : DELTA_CONVEXITY_PRICE_FUNC
'DESCRIPTION   : CONVEXITY ADJUSTMENT; Remember the convexity adjustment gets
'us much closer to the actual value than does the Duration predicted price.
'LIBRARY       : BOND
'GROUP         : DURATION
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************

Function DELTA_CONVEXITY_PRICE_FUNC(ByVal CONVEXITY_VAL As Double, _
ByVal MDURATION_VAL As Double, _
ByVal YIELD1_VAL As Double, _
ByVal YIELD0_VAL As Double)

Dim ADJ_VAL As Double

On Error GoTo ERROR_LABEL
ADJ_VAL = 0.5 * CONVEXITY_VAL * ((YIELD1_VAL - YIELD0_VAL) ^ 2)
DELTA_CONVEXITY_PRICE_FUNC = ADJ_VAL + DELTA_DURATION_PRICE_FUNC(MDURATION_VAL, YIELD1_VAL, YIELD0_VAL)

Exit Function
ERROR_LABEL:
DELTA_CONVEXITY_PRICE_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : DURATION DELTA PRICE
'DESCRIPTION   : DURATION ADJUSTMENT
'LIBRARY       : BOND
'GROUP         : DURATION
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************

Function DELTA_DURATION_PRICE_FUNC(ByVal MDURATION_VAL As Double, _
ByVal YIELD1_VAL As Double, _
ByVal YIELD0_VAL As Double)

On Error GoTo ERROR_LABEL
    
DELTA_DURATION_PRICE_FUNC = -MDURATION_VAL * (YIELD1_VAL - YIELD0_VAL)

Exit Function
ERROR_LABEL:
DELTA_DURATION_PRICE_FUNC = Err.number
End Function
