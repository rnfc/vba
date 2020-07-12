Attribute VB_Name = "FINAN_PORT_MOMENTS_DIETZ_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_RETURNS_SAMPLING_FUNC
'DESCRIPTION   : TWR/MWR Calculations: Time-weighted, money-weighted and "true"
'time-weighted return calculations with user-defined net contributions.
'LIBRARY       : PORTFOLIO
'GROUP         : MOMENTS_DIETZ
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************

Function PORT_TWR_WMR_FUNC(ByRef DATES_RNG As Variant, _
ByRef BMV_RNG As Variant, _
ByRef EMV_RNG As Variant, _
Optional ByVal PREV_EMV_VAL As Double = 0, _
Optional ByVal OUTPUT As Integer = 1)

Dim i As Long
Dim j As Long
Dim NROWS As Long

Dim TEMP1_SUM As Double
Dim TEMP2_SUM As Double
Dim MULT_VAL As Double

Dim DATES_VECTOR As Variant
Dim BMV_VECTOR As Variant
Dim EMV_VECTOR As Variant
Dim TEMP_MATRIX As Variant
Dim TEMP_VECTOR As Variant
Dim CASH_FLOW_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATES_VECTOR = DATES_RNG
If UBound(DATES_VECTOR, 1) = 1 Then
    DATES_VECTOR = MATRIX_TRANSPOSE_FUNC(DATES_VECTOR)
End If
NROWS = UBound(DATES_VECTOR, 1)

BMV_VECTOR = BMV_RNG
If UBound(BMV_VECTOR, 1) = 1 Then
    BMV_VECTOR = MATRIX_TRANSPOSE_FUNC(BMV_VECTOR)
End If
If NROWS <> UBound(BMV_VECTOR, 1) Then: GoTo ERROR_LABEL

EMV_VECTOR = EMV_RNG
If UBound(EMV_VECTOR, 1) = 1 Then
    EMV_VECTOR = MATRIX_TRANSPOSE_FUNC(EMV_VECTOR)
End If
If NROWS <> UBound(EMV_VECTOR, 1) Then: GoTo ERROR_LABEL

TEMP1_SUM = 0: TEMP2_SUM = 0
ReDim TEMP_MATRIX(0 To NROWS, 1 To 8)
'/////////////////////////////////////////////////////////////////////////////////////////////
ReDim CASH_FLOW_VECTOR(1 To NROWS, 1 To 1)
'/////////////////////////////////////////////////////////////////////////////////////////////
TEMP_MATRIX(0, 1) = "DATE"
TEMP_MATRIX(0, 2) = "BMV" 'Beg Market Value
TEMP_MATRIX(0, 3) = "EMV" 'Ending Market Value
TEMP_MATRIX(0, 4) = "NET CONTRIBUTIONS" 'contributions minus withdrawals
'/////////////////////////////////////////////////////////////////////////////////////////////
'Dietz / Modified Dietz Returns
TEMP_MATRIX(0, 5) = "NET CONTRIBUTION WEIGHT"
TEMP_MATRIX(0, 6) = "WEIGHTED NET CONTRIBUTIONS"
'/////////////////////////////////////////////////////////////////////////////////////////////
'Daily TWR
TEMP_MATRIX(0, 7) = "TWR MULTIPLIER"
'/////////////////////////////////////////////////////////////////////////////////////////////
'MWR / IRR
TEMP_MATRIX(0, 8) = "IRR CASH FLOWS"
'/////////////////////////////////////////////////////////////////////////////////////////////

MULT_VAL = 1

i = 1
TEMP_MATRIX(i, 1) = DATES_VECTOR(i, 1)
TEMP_MATRIX(i, 2) = BMV_VECTOR(i, 1)
TEMP_MATRIX(i, 3) = EMV_VECTOR(i, 1)
TEMP_MATRIX(i, 4) = BMV_VECTOR(i, 1) - IIf(PREV_EMV_VAL <> 0, PREV_EMV_VAL, BMV_VECTOR(i, 1))
TEMP1_SUM = TEMP1_SUM + TEMP_MATRIX(i, 4)

TEMP_MATRIX(i, 5) = (NROWS - (i - 1)) / NROWS
TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 4) * TEMP_MATRIX(i, 5)
TEMP2_SUM = TEMP2_SUM + TEMP_MATRIX(i, 6)

TEMP_MATRIX(i, 7) = TEMP_MATRIX(i, 3) / TEMP_MATRIX(i, 2)
MULT_VAL = MULT_VAL * TEMP_MATRIX(i, 7)

TEMP_MATRIX(i, 8) = BMV_VECTOR(i, 1)
CASH_FLOW_VECTOR(i, 1) = TEMP_MATRIX(i, 8)

For i = 2 To NROWS
    TEMP_MATRIX(i, 1) = DATES_VECTOR(i, 1)
    TEMP_MATRIX(i, 2) = BMV_VECTOR(i, 1)
    TEMP_MATRIX(i, 3) = EMV_VECTOR(i, 1)
    TEMP_MATRIX(i, 4) = BMV_VECTOR(i, 1) - EMV_VECTOR(i - 1, 1)
    TEMP1_SUM = TEMP1_SUM + TEMP_MATRIX(i, 4)

    j = i - 1
    TEMP_MATRIX(i, 5) = (NROWS - j) / NROWS
    TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 4) * TEMP_MATRIX(i, 5)
    TEMP2_SUM = TEMP2_SUM + TEMP_MATRIX(i, 6)
    
    TEMP_MATRIX(i, 7) = TEMP_MATRIX(i, 3) / TEMP_MATRIX(i, 2)
    MULT_VAL = MULT_VAL * TEMP_MATRIX(i, 7)

    TEMP_MATRIX(i, 8) = TEMP_MATRIX(i, 4)
    CASH_FLOW_VECTOR(i, 1) = TEMP_MATRIX(i, 8)

Next i
TEMP1_SUM = TEMP1_SUM - TEMP_MATRIX(1, 4)

TEMP_MATRIX(NROWS, 8) = TEMP_MATRIX(NROWS, 8) + -TEMP_MATRIX(NROWS, 3)
CASH_FLOW_VECTOR(NROWS, 1) = TEMP_MATRIX(NROWS, 8)

If OUTPUT = 0 Then
    PORT_TWR_WMR_FUNC = TEMP_MATRIX
    Exit Function
End If
    
ReDim TEMP_VECTOR(1 To 4, 1 To 2)
TEMP_VECTOR(1, 1) = "Dietz Return [%]"
TEMP_VECTOR(1, 2) = 100 * (TEMP_MATRIX(NROWS, 3) - TEMP_MATRIX(1, 2) - TEMP1_SUM) / _
                    (TEMP_MATRIX(1, 2) + 0.5 * TEMP1_SUM)

TEMP_VECTOR(2, 1) = "Modified Dietz Return [%]"
TEMP_VECTOR(2, 2) = 100 * (TEMP_MATRIX(NROWS, 3) - TEMP_MATRIX(1, 2) - TEMP1_SUM) / _
                    (TEMP_MATRIX(1, 2) + TEMP2_SUM)

TEMP_VECTOR(3, 1) = "Daily True TWR [%]"
TEMP_VECTOR(3, 2) = 100 * (MULT_VAL - 1)

TEMP_VECTOR(4, 1) = "Money-Weighted Return (IRR) [%]"
MULT_VAL = IRR_FUNC(CASH_FLOW_VECTOR, DATES_VECTOR, 0.1, 1000, 10 ^ -10)

TEMP_VECTOR(4, 2) = 100 * ((1 + MULT_VAL) ^ ((NROWS - 1) / 365) - 1)

If OUTPUT = 1 Then
    PORT_TWR_WMR_FUNC = TEMP_VECTOR
Else
    PORT_TWR_WMR_FUNC = Array(TEMP_MATRIX, CASH_FLOW_VECTOR)
End If

Exit Function
ERROR_LABEL:
PORT_TWR_WMR_FUNC = Err.number
End Function


'Calculations illustrating the impact of various types of fee charges on Modified Dietz returns.

Function MDR_FEES_FUNC(ByRef BMV_RNG As Variant, _
ByRef EMV_RNG As Variant, _
ByRef NET_CONTRIBUTION_RNG As Variant, _
ByRef WEIGHT_NET_CONTRIBUTION_RNG As Variant, _
ByRef FEES_CHARGED_EXTERNAL_RNG As Variant, _
ByRef WEIGHT_FEES_CHARGED_EXTERNAL_RNG As Variant, _
ByRef FEES_CHARGED_INTERNAL_RNG As Variant, _
ByRef WEIGHT_FEES_CHARGED_INTERNAL_RNG As Variant, _
ByRef FEES_ACCRUED_RNG As Variant)

Dim i As Long
Dim NROWS As Long
Dim TEMP_MATRIX As Variant
Dim BMV_VECTOR As Variant
Dim EMV_VECTOR As Variant
Dim NET_CONTRIBUTION_VECTOR As Variant
Dim WEIGHT_NET_CONTRIBUTION_VECTOR As Variant
Dim FEES_CHARGED_EXTERNAL_VECTOR As Variant
Dim WEIGHT_FEES_CHARGED_EXTERNAL_VECTOR As Variant
Dim FEES_CHARGED_INTERNAL_VECTOR As Variant
Dim WEIGHT_FEES_CHARGED_INTERNAL_VECTOR As Variant
Dim FEES_ACCRUED_VECTOR As Variant
'--------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------------------------------------------
'INPUTS
'1) Beginning Market Value
'2) Ending Market Value
'3) Net Contribution: A value of +10 means that the amount invested was increased by 10. A value of -20 means that an amount of 20 was withdrawn
'4) Weight of Net Contribution: A value of 50% means that the net contributions took place in the middle of the calculation period
'The weight is the "time weight" expressed as % of the period the amount was available for investment purposes in the account.
'5) Fees charged EXTERNAL the account: A value of +5 means that the client paied fees worth 5 from a different account
'6) Weight of fees charged externally*: 0% means that the fee was charged at the end of the period
'7) Fees charged to the account: A value of +5 means that fee charges worth 5 were deducted from the account
'8) Weight of fees charged*: A value of 10% means that deductions took place shortly before the end of the period
'9) Fees accrued during the calculation period: A value of 2 means that fee liabilities of the account increased by 5 during the period
'--------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------------------------------------------
'OUTPUT
'Modified Dietz Return Net-of-Fees
'Modified Dietz Return Gross-of-Fees
'--------------------------------------------------------------------------------------------------------------------------------------------
'Beginning Market Value = 100
'Ending Market Value = 90
'Net Contribution = -20
'Weight of Net Contribution* = 0.5
'Fees charged outside the account = 2
'Weight of fees charged externally* = 0
'Modified Dietz Return Net-of-Fees = 0.0888888888888889
'Fees charged to the account = 2
'Weight of fees charged* = 0.9
'Fees accrued during the calculation period = 1
'Modified Dietz Return Gross-of-Fees = 0.147392290249433
'--------------------------------------------------------------------------------------------------------------------------------------------
On Error GoTo ERROR_LABEL
'--------------------------------------------------------------------------------------------------------------------------------------------
If IsArray(BMV_RNG) = True Then
    BMV_VECTOR = BMV_RNG
    If UBound(BMV_VECTOR, 1) = 1 Then
        BMV_VECTOR = MATRIX_TRANSPOSE_FUNC(BMV_VECTOR)
    End If
Else
    ReDim BMV_VECTOR(1 To 1, 1 To 1)
    BMV_VECTOR(1, 1) = BMV_RNG
End If
NROWS = UBound(BMV_VECTOR, 1)
'--------------------------------------------------------------------------------------------------------------------------------------------
If IsArray(EMV_RNG) = True Then
    EMV_VECTOR = EMV_RNG
    If UBound(EMV_VECTOR, 1) = 1 Then
        EMV_VECTOR = MATRIX_TRANSPOSE_FUNC(EMV_VECTOR)
    End If
Else
    ReDim EMV_VECTOR(1 To 1, 1 To 1)
    EMV_VECTOR(1, 1) = EMV_RNG
End If
If NROWS <> UBound(EMV_VECTOR, 1) Then: GoTo ERROR_LABEL
'--------------------------------------------------------------------------------------------------------------------------------------------
If IsArray(NET_CONTRIBUTION_RNG) = True Then
    NET_CONTRIBUTION_VECTOR = NET_CONTRIBUTION_RNG
    If UBound(NET_CONTRIBUTION_VECTOR, 1) = 1 Then
        NET_CONTRIBUTION_VECTOR = MATRIX_TRANSPOSE_FUNC(NET_CONTRIBUTION_VECTOR)
    End If
Else
    ReDim NET_CONTRIBUTION_VECTOR(1 To 1, 1 To 1)
    NET_CONTRIBUTION_VECTOR(1, 1) = NET_CONTRIBUTION_RNG
End If
If NROWS <> UBound(NET_CONTRIBUTION_VECTOR, 1) Then: GoTo ERROR_LABEL
'--------------------------------------------------------------------------------------------------------------------------------------------
If IsArray(WEIGHT_NET_CONTRIBUTION_RNG) = True Then
    WEIGHT_NET_CONTRIBUTION_VECTOR = WEIGHT_NET_CONTRIBUTION_RNG
    If UBound(WEIGHT_NET_CONTRIBUTION_VECTOR, 1) = 1 Then
        WEIGHT_NET_CONTRIBUTION_VECTOR = MATRIX_TRANSPOSE_FUNC(WEIGHT_NET_CONTRIBUTION_VECTOR)
    End If
Else
    ReDim WEIGHT_NET_CONTRIBUTION_VECTOR(1 To 1, 1 To 1)
    WEIGHT_NET_CONTRIBUTION_VECTOR(1, 1) = WEIGHT_NET_CONTRIBUTION_RNG
End If
If NROWS <> UBound(WEIGHT_NET_CONTRIBUTION_VECTOR, 1) Then: GoTo ERROR_LABEL
'--------------------------------------------------------------------------------------------------------------------------------------------
If IsArray(FEES_CHARGED_EXTERNAL_RNG) = True Then
    FEES_CHARGED_EXTERNAL_VECTOR = FEES_CHARGED_EXTERNAL_RNG
    If UBound(FEES_CHARGED_EXTERNAL_VECTOR, 1) = 1 Then
        FEES_CHARGED_EXTERNAL_VECTOR = MATRIX_TRANSPOSE_FUNC(FEES_CHARGED_EXTERNAL_VECTOR)
    End If
Else
    ReDim FEES_CHARGED_EXTERNAL_VECTOR(1 To 1, 1 To 1)
    FEES_CHARGED_EXTERNAL_VECTOR(1, 1) = FEES_CHARGED_EXTERNAL_RNG
End If
If NROWS <> UBound(FEES_CHARGED_EXTERNAL_VECTOR, 1) Then: GoTo ERROR_LABEL
'--------------------------------------------------------------------------------------------------------------------------------------------
If IsArray(WEIGHT_FEES_CHARGED_EXTERNAL_RNG) = True Then
    WEIGHT_FEES_CHARGED_EXTERNAL_VECTOR = WEIGHT_FEES_CHARGED_EXTERNAL_RNG
    If UBound(WEIGHT_FEES_CHARGED_EXTERNAL_VECTOR, 1) = 1 Then
        WEIGHT_FEES_CHARGED_EXTERNAL_VECTOR = MATRIX_TRANSPOSE_FUNC(WEIGHT_FEES_CHARGED_EXTERNAL_VECTOR)
    End If
Else
    ReDim WEIGHT_FEES_CHARGED_EXTERNAL_VECTOR(1 To 1, 1 To 1)
    WEIGHT_FEES_CHARGED_EXTERNAL_VECTOR(1, 1) = WEIGHT_FEES_CHARGED_EXTERNAL_RNG
End If
If NROWS <> UBound(WEIGHT_FEES_CHARGED_EXTERNAL_VECTOR, 1) Then: GoTo ERROR_LABEL
'--------------------------------------------------------------------------------------------------------------------------------------------
If IsArray(FEES_CHARGED_INTERNAL_RNG) = True Then
    FEES_CHARGED_INTERNAL_VECTOR = FEES_CHARGED_INTERNAL_RNG
    If UBound(FEES_CHARGED_INTERNAL_VECTOR, 1) = 1 Then
        FEES_CHARGED_INTERNAL_VECTOR = MATRIX_TRANSPOSE_FUNC(FEES_CHARGED_INTERNAL_VECTOR)
    End If
Else
    ReDim FEES_CHARGED_INTERNAL_VECTOR(1 To 1, 1 To 1)
    FEES_CHARGED_INTERNAL_VECTOR(1, 1) = FEES_CHARGED_INTERNAL_RNG
End If
If NROWS <> UBound(FEES_CHARGED_INTERNAL_VECTOR, 1) Then: GoTo ERROR_LABEL
'--------------------------------------------------------------------------------------------------------------------------------------------
If IsArray(WEIGHT_FEES_CHARGED_INTERNAL_RNG) = True Then
    WEIGHT_FEES_CHARGED_INTERNAL_VECTOR = WEIGHT_FEES_CHARGED_INTERNAL_RNG
    If UBound(WEIGHT_FEES_CHARGED_INTERNAL_VECTOR, 1) = 1 Then
        WEIGHT_FEES_CHARGED_INTERNAL_VECTOR = MATRIX_TRANSPOSE_FUNC(WEIGHT_FEES_CHARGED_INTERNAL_VECTOR)
    End If
Else
    ReDim WEIGHT_FEES_CHARGED_INTERNAL_VECTOR(1 To 1, 1 To 1)
    WEIGHT_FEES_CHARGED_INTERNAL_VECTOR(1, 1) = WEIGHT_FEES_CHARGED_INTERNAL_RNG
End If
If NROWS <> UBound(WEIGHT_FEES_CHARGED_INTERNAL_VECTOR, 1) Then: GoTo ERROR_LABEL
'--------------------------------------------------------------------------------------------------------------------------------------------
If IsArray(FEES_ACCRUED_RNG) = True Then
    FEES_ACCRUED_VECTOR = FEES_ACCRUED_RNG
    If UBound(FEES_ACCRUED_VECTOR, 1) = 1 Then
        FEES_ACCRUED_VECTOR = MATRIX_TRANSPOSE_FUNC(FEES_ACCRUED_VECTOR)
    End If
Else
    ReDim FEES_ACCRUED_VECTOR(1 To 1, 1 To 1)
    FEES_ACCRUED_VECTOR(1, 1) = FEES_ACCRUED_RNG
End If
If NROWS <> UBound(FEES_ACCRUED_VECTOR, 1) Then: GoTo ERROR_LABEL
'--------------------------------------------------------------------------------------------------------------------------------------------
ReDim TEMP_MATRIX(0 To NROWS, 1 To 11)
'--------------------------------------------------------------------------------------------------------------------------------------------
TEMP_MATRIX(0, 1) = "Beginning Market Value"
TEMP_MATRIX(0, 2) = "Ending Market Value"
TEMP_MATRIX(0, 3) = "Net Contribution"
TEMP_MATRIX(0, 4) = "Weight of Net Contribution"
TEMP_MATRIX(0, 5) = "Fees charged outside the account"
TEMP_MATRIX(0, 6) = "Weight of fees charged externally"
TEMP_MATRIX(0, 7) = "Modified Dietz Return Net-of-Fees"
TEMP_MATRIX(0, 8) = "Fees charged to the account"
TEMP_MATRIX(0, 9) = "Weight of fees charged"
TEMP_MATRIX(0, 10) = "Fees accrued during the calculation period"
TEMP_MATRIX(0, 11) = "Modified Dietz Return Gross-of-Fees"
'--------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------------------------------------------
For i = 1 To NROWS
    TEMP_MATRIX(i, 1) = BMV_VECTOR(i, 1)
    TEMP_MATRIX(i, 2) = EMV_VECTOR(i, 1)
    TEMP_MATRIX(i, 3) = NET_CONTRIBUTION_VECTOR(i, 1)
    TEMP_MATRIX(i, 4) = WEIGHT_NET_CONTRIBUTION_VECTOR(i, 1)
    TEMP_MATRIX(i, 5) = FEES_CHARGED_EXTERNAL_VECTOR(i, 1)
    TEMP_MATRIX(i, 6) = WEIGHT_FEES_CHARGED_EXTERNAL_VECTOR(i, 1)
    TEMP_MATRIX(i, 7) = (TEMP_MATRIX(i, 2) - TEMP_MATRIX(i, 1) - TEMP_MATRIX(i, 3) - TEMP_MATRIX(i, 5)) / (TEMP_MATRIX(i, 1) + TEMP_MATRIX(i, 3) * TEMP_MATRIX(i, 4) + TEMP_MATRIX(i, 5) * TEMP_MATRIX(i, 6))
    TEMP_MATRIX(i, 8) = FEES_CHARGED_INTERNAL_VECTOR(i, 1)
    TEMP_MATRIX(i, 9) = WEIGHT_FEES_CHARGED_INTERNAL_VECTOR(i, 1)
    TEMP_MATRIX(i, 10) = FEES_ACCRUED_VECTOR(i, 1)
    TEMP_MATRIX(i, 11) = (TEMP_MATRIX(i, 2) + TEMP_MATRIX(i, 10) - TEMP_MATRIX(i, 1) - TEMP_MATRIX(i, 3) + TEMP_MATRIX(i, 8)) / (TEMP_MATRIX(i, 1) + TEMP_MATRIX(i, 3) * TEMP_MATRIX(i, 4) + -TEMP_MATRIX(i, 9) * TEMP_MATRIX(i, 8))
Next i
'--------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------------------------------------------
MDR_FEES_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
MDR_FEES_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_RETURN_CONTRIBUTION_FUNC
'DESCRIPTION   : Calculations Modified Dietz Returns on Portfolio & Constituent Level
'LIBRARY       : PORTFOLIO
'GROUP         : MOMENTS_DIETZ
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************

Function PORT_RETURN_CONTRIBUTION_FUNC(ByRef BMV_RNG As Variant, _
ByRef NET_CONTRIBUTION_RNG As Variant, _
ByRef EMV_RNG As Variant, _
Optional ByVal TIME_WEIGHT As Double = 0.5, _
Optional ByVal OUTPUT As Integer = 1)

Dim i As Long
Dim NROWS As Long

Dim BEGINNING_MARKET_VAL As Double
Dim NET_CONTRIBUTION As Double
Dim ENDING_MARKET_VAL As Double
Dim WEIGHT_VAL As Double

Dim WEIGHTED_NET_CONTRIBUTION As Double
Dim MODIFIED_DIETZ_RETURN As Double
Dim AVERAGE_CAPITAL_INVESTED As Double
Dim RETURN_CONTRIBUTION As Double

Dim EMV_VECTOR As Variant
Dim BMV_VECTOR As Variant
Dim NET_CONTRIBUTION_VECTOR As Variant

On Error GoTo ERROR_LABEL

If IsArray(BMV_RNG) = True Then
    BMV_VECTOR = BMV_RNG
    If UBound(BMV_VECTOR, 1) = 1 Then
        BMV_VECTOR = MATRIX_TRANSPOSE_FUNC(BMV_VECTOR)
    End If
Else
    ReDim BMV_VECTOR(1 To 1, 1 To 1)
    BMV_VECTOR(1, 1) = BMV_RNG
End If
NROWS = UBound(BMV_VECTOR, 1)

If IsArray(NET_CONTRIBUTION_RNG) = True Then
    'for example, contribution of 6 takes place mid-period and is invested in A. At 3/4 of
    'the period, 3 is sold from A and invested in B
    NET_CONTRIBUTION_VECTOR = NET_CONTRIBUTION_RNG 'Sum(NET_CONTRIBUTION)
    If UBound(NET_CONTRIBUTION_VECTOR, 1) = 1 Then
        NET_CONTRIBUTION_VECTOR = MATRIX_TRANSPOSE_FUNC(NET_CONTRIBUTION_VECTOR)
    End If
Else
    'for example, net contribution takes place mid-period and is invested immediately in A and B
    ReDim NET_CONTRIBUTION_VECTOR(1 To 1, 1 To 1)
    NET_CONTRIBUTION_VECTOR(1, 1) = NET_CONTRIBUTION_RNG
End If
If NROWS <> UBound(NET_CONTRIBUTION_VECTOR, 1) Then: GoTo ERROR_LABEL

If IsArray(EMV_RNG) = True Then
    EMV_VECTOR = EMV_RNG
    If UBound(EMV_VECTOR, 1) = 1 Then
        EMV_VECTOR = MATRIX_TRANSPOSE_FUNC(EMV_VECTOR)
    End If
Else
    ReDim EMV_VECTOR(1 To 1, 1 To 1)
    EMV_VECTOR(1, 1) = EMV_RNG
End If
If NROWS <> UBound(EMV_VECTOR, 1) Then: GoTo ERROR_LABEL

AVERAGE_CAPITAL_INVESTED = 0
For i = 1 To NROWS
    AVERAGE_CAPITAL_INVESTED = AVERAGE_CAPITAL_INVESTED + BMV_VECTOR(i, 1) + _
                               0.5 * NET_CONTRIBUTION_VECTOR(i, 1)
Next i

BEGINNING_MARKET_VAL = 0
NET_CONTRIBUTION = 0
ENDING_MARKET_VAL = 0
RETURN_CONTRIBUTION = 0
WEIGHTED_NET_CONTRIBUTION = 0

For i = 1 To NROWS
    BEGINNING_MARKET_VAL = BEGINNING_MARKET_VAL + BMV_VECTOR(i, 1)
    NET_CONTRIBUTION = NET_CONTRIBUTION + NET_CONTRIBUTION_VECTOR(i, 1)
    ENDING_MARKET_VAL = ENDING_MARKET_VAL + EMV_VECTOR(i, 1)

    WEIGHTED_NET_CONTRIBUTION = WEIGHTED_NET_CONTRIBUTION + NET_CONTRIBUTION_VECTOR(i, 1) * TIME_WEIGHT
    MODIFIED_DIETZ_RETURN = (EMV_VECTOR(i, 1) - BMV_VECTOR(i, 1) - _
                            NET_CONTRIBUTION_VECTOR(i, 1)) / (BMV_VECTOR(i, 1) + _
                            (NET_CONTRIBUTION_VECTOR(i, 1) * TIME_WEIGHT))

    WEIGHT_VAL = (BMV_VECTOR(i, 1) + 0.5 * NET_CONTRIBUTION_VECTOR(i, 1)) / AVERAGE_CAPITAL_INVESTED
    
    RETURN_CONTRIBUTION = RETURN_CONTRIBUTION + MODIFIED_DIETZ_RETURN * WEIGHT_VAL
Next i

'------------------------------------------------------------------------------------------
Select Case OUTPUT
'------------------------------------------------------------------------------------------
Case 0
'------------------------------------------------------------------------------------------
    PORT_RETURN_CONTRIBUTION_FUNC = RETURN_CONTRIBUTION
'------------------------------------------------------------------------------------------
Case Else
'------------------------------------------------------------------------------------------
    
    MODIFIED_DIETZ_RETURN = (ENDING_MARKET_VAL - BEGINNING_MARKET_VAL - _
                             NET_CONTRIBUTION) / (BEGINNING_MARKET_VAL + _
                             (NET_CONTRIBUTION * TIME_WEIGHT))
    
    ReDim TEMP_VECTOR(1 To 9, 1 To 2)
    TEMP_VECTOR(1, 1) = "BEGINNING MARKET VALUE"
    TEMP_VECTOR(1, 2) = BEGINNING_MARKET_VAL
    
    TEMP_VECTOR(2, 1) = "NET CONTRIBUTION"
    TEMP_VECTOR(2, 2) = NET_CONTRIBUTION
    
    TEMP_VECTOR(3, 1) = "TIME WEIGHT"
    TEMP_VECTOR(3, 2) = TIME_WEIGHT
    
    TEMP_VECTOR(4, 1) = "WEIGHTED NET CONTRIBUTION"
    TEMP_VECTOR(4, 2) = WEIGHTED_NET_CONTRIBUTION
    
    TEMP_VECTOR(5, 1) = "ENDING MARKET VALUE"
    TEMP_VECTOR(5, 2) = ENDING_MARKET_VAL
    
    TEMP_VECTOR(6, 1) = "MODIFIED-DIETZ RETURN"
    TEMP_VECTOR(6, 2) = MODIFIED_DIETZ_RETURN
    
    TEMP_VECTOR(7, 1) = "AVERAGE CAPITAL INVESTED"
    TEMP_VECTOR(7, 2) = AVERAGE_CAPITAL_INVESTED
    
    TEMP_VECTOR(8, 1) = "WEIGHT"
    TEMP_VECTOR(8, 2) = 1
    
    TEMP_VECTOR(9, 1) = "RETURN CONTRIBUTION"
    TEMP_VECTOR(9, 2) = RETURN_CONTRIBUTION

    PORT_RETURN_CONTRIBUTION_FUNC = TEMP_VECTOR
'------------------------------------------------------------------------------------------
End Select
'------------------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
PORT_RETURN_CONTRIBUTION_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CONSTITUENT_RETURN_CONTRIBUTION_FUNC
'DESCRIPTION   : Calculations Modified Dietz Returns on Constituent Level
'LIBRARY       : PORTFOLIO
'GROUP         : MOMENTS_DIETZ
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************

Function CONSTITUENT_RETURN_CONTRIBUTION_FUNC(ByVal BEGINNING_MARKET_VAL As Double, _
ByVal NET_CONTRIBUTION_RNG As Variant, _
ByVal TIME_WEIGHTS_RNG As Variant, _
ByVal ENDING_MARKET_VAL As Double, _
ByVal WEIGHT_VAL As Double)

Dim TEMP_ARR As Variant
Dim NET_CONTRIBUTION As Double
Dim WEIGHTED_NET_CONTRIBUTION As Double
Dim MODIFIED_DIETZ_RETURN As Double
Dim AVERAGE_CAPITAL_INVESTED As Double

On Error GoTo ERROR_LABEL

TEMP_ARR = CONSTITUENT_WEIGHTED_NET_CONTRIBUTION_FUNC(NET_CONTRIBUTION_RNG, TIME_WEIGHTS_RNG)
If IsArray(TEMP_ARR) = False Then: GoTo ERROR_LABEL

NET_CONTRIBUTION = TEMP_ARR(LBound(TEMP_ARR))
WEIGHTED_NET_CONTRIBUTION = TEMP_ARR(UBound(TEMP_ARR))

MODIFIED_DIETZ_RETURN = (ENDING_MARKET_VAL - BEGINNING_MARKET_VAL - _
                        NET_CONTRIBUTION) / (BEGINNING_MARKET_VAL + _
                        WEIGHTED_NET_CONTRIBUTION)
AVERAGE_CAPITAL_INVESTED = BEGINNING_MARKET_VAL + 0.5 * NET_CONTRIBUTION

CONSTITUENT_RETURN_CONTRIBUTION_FUNC = MODIFIED_DIETZ_RETURN * WEIGHT_VAL

Exit Function
ERROR_LABEL:
CONSTITUENT_RETURN_CONTRIBUTION_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CONSTITUENT_WEIGHTED_NET_CONTRIBUTION_FUNC
'DESCRIPTION   : Disaggregating Portfolio Modified-Dietz Returns
'weighted net contribution at constituent
'LIBRARY       : PORTFOLIO
'GROUP         : MOMENTS_DIETZ
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************

Function CONSTITUENT_WEIGHTED_NET_CONTRIBUTION_FUNC( _
ByRef NET_CONTRIBUTION_RNG As Variant, _
ByRef TIME_WEIGHTS_RNG As Variant)

Dim i As Long
Dim NROWS As Long
Dim TEMP1_SUM As Double 'Net Contribution
Dim TEMP2_SUM As Double 'Weighted Net Contribution

Dim TIME_WEIGHTS_VECTOR As Variant
Dim NET_CONTRIBUTION_VECTOR As Variant

On Error GoTo ERROR_LABEL

If IsArray(NET_CONTRIBUTION_RNG) = True Then
    'for example, contribution of 6 takes place mid-period and is invested in A. At 3/4 of
    'the period, 3 is sold from A and invested in B
    NET_CONTRIBUTION_VECTOR = NET_CONTRIBUTION_RNG 'Sum(NET_CONTRIBUTION)
    If UBound(NET_CONTRIBUTION_VECTOR, 1) = 1 Then
        NET_CONTRIBUTION_VECTOR = MATRIX_TRANSPOSE_FUNC(NET_CONTRIBUTION_VECTOR)
    End If
Else
    'for example, net contribution takes place mid-period and is invested immediately in A and B
    ReDim NET_CONTRIBUTION_VECTOR(1 To 1, 1 To 1)
    NET_CONTRIBUTION_VECTOR(1, 1) = NET_CONTRIBUTION_RNG
End If
NROWS = UBound(NET_CONTRIBUTION_VECTOR, 1)

If IsArray(TIME_WEIGHTS_RNG) = True Then
    TIME_WEIGHTS_VECTOR = TIME_WEIGHTS_RNG
    If UBound(TIME_WEIGHTS_VECTOR, 1) = 1 Then
        TIME_WEIGHTS_VECTOR = MATRIX_TRANSPOSE_FUNC(TIME_WEIGHTS_VECTOR)
    End If
Else
    ReDim TIME_WEIGHTS_VECTOR(1 To 1, 1 To 1)
    TIME_WEIGHTS_VECTOR(1, 1) = TIME_WEIGHTS_RNG
End If
If NROWS <> UBound(TIME_WEIGHTS_VECTOR, 1) Then: GoTo ERROR_LABEL

TEMP1_SUM = 0: TEMP2_SUM = 0
For i = 1 To NROWS
    TEMP1_SUM = TEMP1_SUM + NET_CONTRIBUTION_VECTOR(i, 1)
    TEMP2_SUM = TEMP2_SUM + NET_CONTRIBUTION_VECTOR(i, 1) * TIME_WEIGHTS_VECTOR(i, 1)
Next i

CONSTITUENT_WEIGHTED_NET_CONTRIBUTION_FUNC = Array(TEMP1_SUM, TEMP2_SUM)

Exit Function
ERROR_LABEL:
CONSTITUENT_WEIGHTED_NET_CONTRIBUTION_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : RNG_PORT_TIME_WEIGHTED_RETURNS_FUNC
'DESCRIPTION   : Calculating composite returns from time-weighted portfolio returns
'LIBRARY       : PORTFOLIO
'GROUP         : RISK
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************
'RNG_PORT_WEIGHTED_CASH_FLOW_FUNC
Function RNG_PORT_TIME_WEIGHTED_RETURNS_FUNC(ByRef DST_RNG As Excel.Range, _
ByVal NASSETS As Long, _
Optional ByVal PERIODS As Long = 3, _
Optional ByVal COUNT_BASIS As Double = 30, _
Optional ByVal ADD_RNG_NAME As Boolean = False)

Dim i As Long
Dim j As Long
Dim m As Long

Dim DATA_POS_RNG As Excel.Range
Dim DATA_RNG As Excel.Range

Dim CASH_POS_RNG As Excel.Range
Dim CASH_RNG As Excel.Range

Dim MV_POS_RNG As Excel.Range
Dim MV_RNG As Excel.Range

Dim TIME_POS_RNG As Excel.Range
Dim TIME_RNG As Excel.Range

On Error GoTo ERROR_LABEL

RNG_PORT_TIME_WEIGHTED_RETURNS_FUNC = False
m = 4
If (NASSETS < 2) Or (PERIODS < 1) Then: GoTo ERROR_LABEL

'---------------------------------------------------------------------------
'------------------------FIRT PASS: PORT. DATA -----------------------------
'---------------------------------------------------------------------------

Set DATA_POS_RNG = DST_RNG
DATA_POS_RNG.Offset(0, 0).value = "PORTFOLIO DATA"
DATA_POS_RNG.Offset(0, 0).Font.Bold = True

DATA_POS_RNG.Offset(0, 3 + PERIODS).value = "Cash Flow"
DATA_POS_RNG.Offset(0, 3 + PERIODS).Font.Bold = True

DATA_POS_RNG.Offset(0, 4 + PERIODS).value = "Total Cash Flow"
DATA_POS_RNG.Offset(0, 4 + PERIODS).Font.Bold = True

With DATA_POS_RNG
   Set DATA_RNG = Range(.Offset(NASSETS + 1, 1), .Offset(2, 4 + PERIODS))
   If ADD_RNG_NAME = True Then: DATA_RNG.name = "DATA_STAT"
    
    .Offset(0, 1).value = "BEG MV"
    .Offset(0, 1).Font.Bold = True
    
    .Offset(0, 2).value = "END MV"
    .Offset(0, 2).Font.Bold = True
    
    .Offset(1, 0).value = "Period"
    .Offset(1, 0).Font.Bold = True

    For i = 1 To NASSETS
      With .Offset(i + 1, 0)
         .value = "Asset " & CStr(i)
         .Font.ColorIndex = 3
      End With
    Next i
End With

With DATA_RNG
    .Columns(1).Cells(NASSETS + 1, 0).value = "Composite Total"
    .Columns(1).Cells(NASSETS + 1, 0).Font.Bold = True
    .Columns(1).value = 0
    .Columns(1).Font.ColorIndex = 5
    .Columns(1).Cells(NASSETS + 1).formula = "=SUM(" & .Columns(1).Address & ")"
    
    .Columns(2).value = 0
    .Columns(2).Font.ColorIndex = 5
    .Columns(2).Cells(NASSETS + 1).formula = "=SUM(" & .Columns(2).Address & ")"
    
    Range(.Columns(4), .Columns(4 + PERIODS - 1)).value = 0
    Range(.Columns(4), .Columns(4 + PERIODS - 1)).Font.ColorIndex = 5
    
        For i = 1 To PERIODS + 1
            With .Columns(4).Cells(NASSETS + 1)
                .Offset(0, i - 1).formula = "=SUM(" & _
                DATA_RNG.Columns(3 + i).Address & ")"
            End With
        Next i
        For i = 1 To PERIODS
            With .Columns(4).Cells(0)
                .Offset(0, i - 1).value = 0
                .Offset(0, i - 1).Font.ColorIndex = 3
            End With
        Next i
        
        For i = 1 To NASSETS
            With .Columns(PERIODS + 4)
                .Cells(i).formula = "=SUM(" & _
                Range(DATA_RNG.Columns(4), _
                DATA_RNG.Columns(4 + PERIODS - 1)).Rows(i).Address & ")"
            End With
        Next i
    
End With

'---------------------------------------------------------------------------
'-------------------SECOND PASS: TIME-WEIGHTED CASH FLOWS-------------------
'---------------------------------------------------------------------------

Set CASH_POS_RNG = DATA_POS_RNG.Offset(NASSETS + 1 + m, 0)
CASH_POS_RNG.Offset(0, 0).value = "CALCULATION OF TIME-WEIGHTED CASH FLOWS"
CASH_POS_RNG.Offset(0, 0).Font.Bold = True

CASH_POS_RNG.Offset(0, 3 + PERIODS).value = "Time-Weighted Cash Flow"
CASH_POS_RNG.Offset(0, 3 + PERIODS).Font.Bold = True

CASH_POS_RNG.Offset(0, 4 + PERIODS).value = "Total Time-Weighted Cash Flow"
CASH_POS_RNG.Offset(0, 4 + PERIODS).Font.Bold = True

With CASH_POS_RNG
   Set CASH_RNG = Range(.Offset(NASSETS + 1, 1), .Offset(2, 4 + PERIODS))
   If ADD_RNG_NAME = True Then: CASH_RNG.name = _
        "CAHS_STAT"

    .Offset(1, 0).value = "Period"
    .Offset(1, 0).Font.Bold = True

    For i = 1 To NASSETS
      With .Offset(i + 1, 0)
         .formula = "=" & DATA_POS_RNG.Offset(i + 1, 0).Address
      End With
    Next i
End With

With CASH_RNG
    .Columns(1).Cells(NASSETS + 1, 0).value = "Composite Total"
    .Columns(1).Cells(NASSETS + 1, 0).Font.Bold = True
    
        For j = 1 To PERIODS
            For i = 1 To NASSETS
                With .Columns(3 + j).Cells(1)
                    .Offset(i - 1, 0).formula = "=" & _
                    CASH_RNG.Columns(3 + j).Cells(0).Address & _
                    "*" & DATA_RNG.Columns(3 + j).Cells(i).Address
                End With
            Next i
        Next j
    
        For i = 1 To PERIODS + 1
            With .Columns(4).Cells(NASSETS + 1)
                .Offset(0, i - 1).formula = "=SUM(" & _
                CASH_RNG.Columns(3 + i).Address & ")"
            End With
        Next i
        
        For i = 1 To PERIODS
            With .Columns(4).Cells(0)
                .Offset(0, i - 1).formula = "=1-" & _
                DATA_RNG.Columns(4).Cells(0).Offset(0, i - 1).Address & _
                "/" & COUNT_BASIS
                .Offset(0, i - 1).Font.ColorIndex = 3
            End With
        Next i
        
        For i = 1 To NASSETS
            With .Columns(PERIODS + 4)
                .Cells(i).formula = "=SUM(" & _
                Range(CASH_RNG.Columns(4), _
                CASH_RNG.Columns(4 + PERIODS - 1)).Rows(i).Address & ")"
            End With
        Next i
    
End With

'---------------------------------------------------------------------------
'-------------------THIRD PASS:ADJUSTED BEGINNING MARKET VALUE--------------
'---------------------------------------------------------------------------

Set MV_POS_RNG = CASH_POS_RNG.Offset(NASSETS + 1 + m, 0)
MV_POS_RNG.Offset(0, 0).value = "CALCULATION OF ADJUSTED BEGINNING MARKET VALUE"
MV_POS_RNG.Offset(0, 0).Font.Bold = True

MV_POS_RNG.Offset(0, 1).value = "Beginning Market Value"
MV_POS_RNG.Offset(0, 1).Font.Bold = True

MV_POS_RNG.Offset(0, 2).value = "Total Time-Weighted Cash Flow"
MV_POS_RNG.Offset(0, 2).Font.Bold = True

MV_POS_RNG.Offset(0, 3).value = "Adjusted Beginning Market Value"
MV_POS_RNG.Offset(0, 3).Font.Bold = True

With MV_POS_RNG
   Set MV_RNG = Range(.Offset(NASSETS + 1, 1), .Offset(2, 3))
   If ADD_RNG_NAME = True Then: MV_RNG.name = _
        "MV_STAT"

    For i = 1 To NASSETS
      With .Offset(i + 1, 0)
         .formula = "=" & DATA_POS_RNG.Offset(i + 1, 0).Address
      End With
      With .Offset(i + 1, 1)
         .formula = "=" & DATA_POS_RNG.Offset(i + 1, 1).Address
      End With
      With .Offset(i + 1, 2)
         .formula = "=" & CASH_POS_RNG.Offset(i + 1, 4 + PERIODS).Address
      End With
      With .Offset(i + 1, 3)
         .formula = "=" & MV_POS_RNG.Offset(i + 1, 1).Address & _
         "+" & MV_POS_RNG.Offset(i + 1, 2).Address
      End With
    Next i
    .Offset(NASSETS + 2, 0).value = "Composite Total"
    .Offset(NASSETS + 2, 0).Font.Bold = True
    
    .Offset(NASSETS + 2, 1).formula = "=SUM(" & MV_RNG.Columns(1).Address & ")"
    .Offset(NASSETS + 2, 2).formula = "=SUM(" & MV_RNG.Columns(2).Address & ")"
    .Offset(NASSETS + 2, 3).formula = "=SUM(" & MV_RNG.Columns(3).Address & ")"
End With

'---------------------------------------------------------------------------
'---------------------FORTH PASS:TIME-WEIGHTED RETURNS----------------------
'---------------------------------------------------------------------------

Set TIME_POS_RNG = MV_POS_RNG.Offset(NASSETS + 1 + m, 0)
TIME_POS_RNG.Offset(0, 0).value = "CALCULATION OF TIME-WEIGHTED RETURNS"
TIME_POS_RNG.Offset(0, 0).Font.Bold = True

TIME_POS_RNG.Offset(0, 1).value = "Time-Weighted Return"
TIME_POS_RNG.Offset(0, 1).Font.Bold = True

TIME_POS_RNG.Offset(0, 2).value = "Weight Based on Beginning Market Value"
TIME_POS_RNG.Offset(0, 2).Font.Bold = True

TIME_POS_RNG.Offset(0, 3).value = "Weight based on Adjusted Beginning Market Value"
TIME_POS_RNG.Offset(0, 3).Font.Bold = True

TIME_POS_RNG.Offset(0, 4).value = "Asset-Weighted Composite Return"
TIME_POS_RNG.Offset(0, 4).Font.Bold = True

TIME_POS_RNG.Offset(0, 5).value = "Asset-Weighted and Cash-Flow-Weighted Returns"
TIME_POS_RNG.Offset(0, 5).Font.Bold = True

With TIME_POS_RNG
   Set TIME_RNG = Range(.Offset(NASSETS + 1, 1), .Offset(2, 5))
   If ADD_RNG_NAME = True Then: TIME_RNG.name = _
        "TIME_STAT"

    For i = 1 To NASSETS
      With .Offset(i + 1, 0)
         .formula = "=" & DATA_POS_RNG.Offset(i + 1, 0).Address
      End With
      With .Offset(i + 1, 1)
         .formula = "=(" & DATA_POS_RNG.Offset(i + 1, 2).Address & "-" & _
                    DATA_POS_RNG.Offset(i + 1, 1).Address & "-" & _
                    DATA_POS_RNG.Offset(i + 1, 4 + PERIODS).Address & ")/(" & _
                    DATA_POS_RNG.Offset(i + 1, 1).Address & "+" & _
                    CASH_POS_RNG.Offset(i + 1, 4 + PERIODS).Address & ")"
      End With
      With .Offset(i + 1, 2)
         .formula = "=" & DATA_POS_RNG.Offset(i + 1, 1).Address & "/" & _
                    DATA_POS_RNG.Offset(NASSETS + 2, 1).Address
      End With
      With .Offset(i + 1, 3)
         .formula = "=" & MV_POS_RNG.Offset(i + 1, 3).Address & "/" & _
                    MV_POS_RNG.Offset(NASSETS + 2, 3).Address
      End With
      With .Offset(i + 1, 4)
         .formula = "=" & TIME_POS_RNG.Offset(i + 1, 1).Address & _
         "*" & TIME_POS_RNG.Offset(i + 1, 2).Address
      End With
      With .Offset(i + 1, 5)
         .formula = "=" & TIME_POS_RNG.Offset(i + 1, 1).Address & _
         "*" & TIME_POS_RNG.Offset(i + 1, 3).Address
      End With
    
    Next i
    .Offset(NASSETS + 2, 0).value = "Composite Total"
    .Offset(NASSETS + 2, 0).Font.Bold = True
    
    .Offset(NASSETS + 2, 1).formula = "=(" & _
                    DATA_POS_RNG.Offset(NASSETS + 2, 2).Address & "-" & _
                    DATA_POS_RNG.Offset(NASSETS + 2, 1).Address & "-" & _
                    DATA_POS_RNG.Offset(NASSETS + 2, 4 + PERIODS).Address & ")/(" & _
                    DATA_POS_RNG.Offset(NASSETS + 2, 1).Address & "+" & _
                    CASH_POS_RNG.Offset(NASSETS + 2, 4 + PERIODS).Address & ")"

    .Offset(NASSETS + 2, 2).formula = "=SUM(" & TIME_RNG.Columns(2).Address & ")"
    .Offset(NASSETS + 2, 3).formula = "=SUM(" & TIME_RNG.Columns(3).Address & ")"
    .Offset(NASSETS + 2, 4).formula = "=SUM(" & TIME_RNG.Columns(4).Address & ")"
    .Offset(NASSETS + 2, 5).formula = "=SUM(" & TIME_RNG.Columns(5).Address & ")"
End With

RNG_PORT_TIME_WEIGHTED_RETURNS_FUNC = True

Exit Function
ERROR_LABEL:
RNG_PORT_TIME_WEIGHTED_RETURNS_FUNC = False
End Function

