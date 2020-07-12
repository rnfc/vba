Attribute VB_Name = "FINAN_ASSET_MOMENTS_DIETZ_LIBR"



Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
Private Const PUB_EPSILON As Double = 2 ^ 52

'MWR/TWR/MD/D/ROI Return Calculations: Calendar year calculations for
'money-weighted, time-weighted, Dietz and modified Dietz returns plus
'return on investment (ROI). The calendar year returns are compared
'to annualized monthly return calculations, which allows an assessment of
'various approximations used in investment management.

Function ASSET_MTMDR_FUNC(ByRef DATE_RNG As Variant, _
ByRef EMV_RNG As Variant, _
ByRef CONTRIBUTION_RNG As Variant, _
Optional ByVal INITIAL_PORT_VALUE As Double = 100, _
Optional ByVal DAYS_PER_YEAR As Long = 366, _
Optional ByVal OUTPUT As Integer = 0, _
Optional ByVal FACTOR_VAL As Double = 100)

'Net Contributions are defined as contributions minus withdrawals


Dim h() As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long
Dim m As Long
Dim n As Long

Dim NROWS As Long

Dim DATE0_VAL As Date
Dim DATE1_VAL As Date
Dim DATE2_VAL As Date

Dim WITHDRAWALS_VAL As Double
Dim CONTRIBUTIONS_VAL As Double
Dim NET_CONTRIBUTIONS_VAL As Double
Dim WEIGHTED_NET_CONTRIBUTIONS_VAL As Double

Dim DATA_MATRIX As Variant
Dim SUMMARY_VECTOR As Variant

Dim EMV_VECTOR As Variant
Dim DATE_VECTOR  As Variant
Dim CONTRIBUTION_VECTOR As Variant

Dim XIRR_VAL As Variant

Dim DATE2_VECTOR() As Date
Dim DATA2_VECTOR() As Double

On Error GoTo ERROR_LABEL

DATE_VECTOR = DATE_RNG
If UBound(DATE_VECTOR, 1) = 1 Then
    DATE_VECTOR = MATRIX_TRANSPOSE_FUNC(DATE_VECTOR)
End If
NROWS = UBound(DATE_VECTOR, 1)

EMV_VECTOR = EMV_RNG
If UBound(EMV_VECTOR, 1) = 1 Then
    EMV_VECTOR = MATRIX_TRANSPOSE_FUNC(EMV_VECTOR)
End If
If NROWS <> UBound(EMV_VECTOR, 1) Then: GoTo ERROR_LABEL

CONTRIBUTION_VECTOR = CONTRIBUTION_RNG
If UBound(CONTRIBUTION_VECTOR, 1) = 1 Then
    CONTRIBUTION_VECTOR = MATRIX_TRANSPOSE_FUNC(CONTRIBUTION_VECTOR)
End If
If NROWS <> UBound(CONTRIBUTION_VECTOR, 1) Then: GoTo ERROR_LABEL

'-----------------------------------------------------------------------------------
ReDim SUMMARY_VECTOR(0 To 1, 1 To 10)
'Calendar Year Return Calculations & Annualized Monthly Return Calculations
'-----------------------------------------------------------------------------------
SUMMARY_VECTOR(0, 1) = "ANNUAL DIETZ RETURN [%]"
SUMMARY_VECTOR(0, 2) = "ANNUAL MODIFIED DIETZ RETURN [%]" '...is an approximation for the Annual True MWR
SUMMARY_VECTOR(0, 3) = "ANNUAL TRUE TWR (DAILY CHAIN-LINKED) [%]"
SUMMARY_VECTOR(0, 4) = "ANNUAL TRUE MWR (DAILY XIRR) [%]"
SUMMARY_VECTOR(0, 5) = "ANNUAL RETURN ON INVESTMENT [%]" '...takes into account changes in capital invested, but not the timing of the net contributions

SUMMARY_VECTOR(0, 6) = "CHAIN-LINKED MONTHLY DIETZ RETURN [%]"
SUMMARY_VECTOR(0, 7) = "CHAIN-LINKED MONTHLY MODIFIED DIETZ RETURN [%]" '...is the approximation to Annual True TWR promoted by GIPS
SUMMARY_VECTOR(0, 8) = "CHAIN-LINKED MONTHLY TRUE TWR [%]" '...is exactly the same as the Annual True TWR
SUMMARY_VECTOR(0, 9) = "CHAIN-LINKED MONTHLY TRUE MWR (XIRR) [%]" '...is an approximation to Annual True TWR (and not Annual True MWR)
SUMMARY_VECTOR(0, 10) = "CHAIN-LINKED ANNUAL RETURN ON INVESTMENT [%]" '...is a surprisingly good approximation for the True TWR
'-----------------------------------------------------------------------------------
ReDim DATA_MATRIX(0 To NROWS, 1 To 15)
'-----------------------------------------------------------------------------------
DATA_MATRIX(0, 1) = "DATE"
DATA_MATRIX(0, 2) = "ANNUAL TIME POINT (BEGINNING OF DAY)"
DATA_MATRIX(0, 3) = "MONTHLY TIME POINT (BEGINNING OF DAY)"
DATA_MATRIX(0, 4) = "DAILY RETURNS"
DATA_MATRIX(0, 5) = "NET CONTRIBUTIONS**"
DATA_MATRIX(0, 6) = "DAILY BMV"
DATA_MATRIX(0, 7) = "DAILY EMV"
'-----------------------------------------------------------------------------------
DATA_MATRIX(0, 8) = "MONTHLY DIETZ / MODIFIED DIETZ RETURNS: MONTHLY NET CONTRIBUTION WEIGHT"
DATA_MATRIX(0, 9) = "MONTHLY DIETZ / MODIFIED DIETZ RETURNS: MONTHLY WEIGHTED NET CONTRIBUTIONS**"
DATA_MATRIX(0, 10) = "ANNUAL DIETZ / MODIFIED DIETZ RETURNS: ANNUAL NET CONTRIBUTION WEIGHT"
DATA_MATRIX(0, 11) = "ANNUAL DIETZ / MODIFIED DIETZ RETURNS: ANNUAL WEIGHTED NET CONTRIBUTIONS**"
DATA_MATRIX(0, 12) = "DAILY TRUE TWR MULTIPLIER"
DATA_MATRIX(0, 13) = "NET CONTRIBUTION AS % OF MV"
DATA_MATRIX(0, 14) = "ANNUAL XIRR CASH FLOWS"
DATA_MATRIX(0, 15) = "MONTHLY XIRR CF"
'-----------------------------------------------------------------------------------

WITHDRAWALS_VAL = 0: CONTRIBUTIONS_VAL = 0
NET_CONTRIBUTIONS_VAL = 0: WEIGHTED_NET_CONTRIBUTIONS_VAL = 0

SUMMARY_VECTOR(1, 3) = 1
ReDim DATE2_VECTOR(1 To NROWS, 1 To 1)
ReDim DATA2_VECTOR(1 To NROWS, 1 To 1)

DATE0_VAL = DATE_VECTOR(1, 1)
k = Year(DATE0_VAL)
l = 1
ReDim h(1 To 2, 1 To l)
h(1, l) = 1
h(2, l) = Month(DATE0_VAL)

For i = 1 To NROWS
    DATA_MATRIX(i, 1) = DATE_VECTOR(i, 1)
    DATA_MATRIX(i, 2) = DATA_MATRIX(i, 1) - DATE0_VAL
    
    j = Day(DATA_MATRIX(i, 1))
    DATA_MATRIX(i, 3) = j - 1
    DATA_MATRIX(i, 5) = CONTRIBUTION_VECTOR(i, 1)
    
    NET_CONTRIBUTIONS_VAL = NET_CONTRIBUTIONS_VAL + DATA_MATRIX(i, 5)
    If DATA_MATRIX(i, 5) > 0 Then
        CONTRIBUTIONS_VAL = CONTRIBUTIONS_VAL + DATA_MATRIX(i, 5)
    ElseIf DATA_MATRIX(i, 5) < 0 Then
        WITHDRAWALS_VAL = WITHDRAWALS_VAL + DATA_MATRIX(i, 5)
    End If
    
    DATA_MATRIX(i, 7) = EMV_VECTOR(i, 1)
    If i <> 1 Then
        DATA_MATRIX(i, 6) = DATA_MATRIX(i - 1, 7) + DATA_MATRIX(i, 5)
    Else
        DATA_MATRIX(i, 6) = INITIAL_PORT_VALUE
    End If
    
    DATA_MATRIX(i, 4) = FACTOR_VAL * (DATA_MATRIX(i, 7) / DATA_MATRIX(i, 6) - 1)
    j = Month(DATA_MATRIX(i, 1))
    If j <> h(2, l) Then
        l = l + 1
        ReDim Preserve h(1 To 2, 1 To l)
        h(1, l) = i
        h(2, l) = j
    End If
    
    DATE1_VAL = DateSerial(k, j, 1)
    DATE2_VAL = DateSerial(k, j + 1, 1)
    DATA_MATRIX(i, 8) = ((DATE2_VAL - DATE1_VAL) - DATA_MATRIX(i, 3)) / (DATE2_VAL - DATE1_VAL)
    DATA_MATRIX(i, 9) = DATA_MATRIX(i, 8) * DATA_MATRIX(i, 5)
    DATA_MATRIX(i, 10) = (DAYS_PER_YEAR - DATA_MATRIX(i, 2)) / DAYS_PER_YEAR
    
    DATA_MATRIX(i, 11) = DATA_MATRIX(i, 5) * DATA_MATRIX(i, 10)
    WEIGHTED_NET_CONTRIBUTIONS_VAL = WEIGHTED_NET_CONTRIBUTIONS_VAL + DATA_MATRIX(i, 11)
    
    DATA_MATRIX(i, 12) = DATA_MATRIX(i, 7) / DATA_MATRIX(i, 6)
    SUMMARY_VECTOR(1, 3) = SUMMARY_VECTOR(1, 3) * DATA_MATRIX(i, 12)
    
    DATA_MATRIX(i, 13) = FACTOR_VAL * DATA_MATRIX(i, 5) / DATA_MATRIX(i, 7)
    
    If i <> 1 Then
        If DATA_MATRIX(i, 2) + 1 = DAYS_PER_YEAR Then
            DATA_MATRIX(i, 14) = DATA_MATRIX(i, 5) - DATA_MATRIX(i, 7)
        Else
            DATA_MATRIX(i, 14) = DATA_MATRIX(i, 5) + 0
        End If
    Else
        DATA_MATRIX(i, 14) = DATA_MATRIX(i, 5) + INITIAL_PORT_VALUE
    End If
    
    DATE2_VECTOR(i, 1) = DATA_MATRIX(i, 1)
    DATA2_VECTOR(i, 1) = DATA_MATRIX(i, 14)
        
    If DATA_MATRIX(i, 3) + 1 = (DATE2_VAL - DATE1_VAL) Then
        If DATA_MATRIX(i, 3) = 0 Then
            DATA_MATRIX(i, 15) = DATA_MATRIX(i, 5) - DATA_MATRIX(i, 7) + DATA_MATRIX(i, 6)
        Else
            DATA_MATRIX(i, 15) = DATA_MATRIX(i, 5) - DATA_MATRIX(i, 7) + 0
        End If
    Else
        If DATA_MATRIX(i, 3) = 0 Then
            DATA_MATRIX(i, 15) = DATA_MATRIX(i, 5) + 0 + DATA_MATRIX(i, 6)
        Else
            DATA_MATRIX(i, 15) = DATA_MATRIX(i, 5) + 0 + 0
        End If
    End If
    
    'If DATA_MATRIX(i, 5) = 0 Then: DATA_MATRIX(i, 5) = ""
    'If DATA_MATRIX(i, 9) = 0 Then: DATA_MATRIX(i, 9) = ""
    'If DATA_MATRIX(i, 11) = 0 Then: DATA_MATRIX(i, 11) = ""
    'If DATA_MATRIX(i, 13) = 0 Then: DATA_MATRIX(i, 13) = ""
    
Next i

SUMMARY_VECTOR(1, 1) = FACTOR_VAL * (DATA_MATRIX(NROWS, 7) - DATA_MATRIX(1, 6) - NET_CONTRIBUTIONS_VAL) / (DATA_MATRIX(1, 6) + 0.5 * NET_CONTRIBUTIONS_VAL)
SUMMARY_VECTOR(1, 2) = FACTOR_VAL * (DATA_MATRIX(NROWS, 7) - DATA_MATRIX(1, 6) - NET_CONTRIBUTIONS_VAL) / (DATA_MATRIX(1, 6) + WEIGHTED_NET_CONTRIBUTIONS_VAL)
SUMMARY_VECTOR(1, 3) = FACTOR_VAL * (SUMMARY_VECTOR(1, 3) - 1)
XIRR_VAL = IRR_FUNC(DATA2_VECTOR, DATE2_VECTOR, FACTOR_VAL, 1000, 10 ^ -10)
If XIRR_VAL <> PUB_EPSILON Then
    SUMMARY_VECTOR(1, 4) = FACTOR_VAL * ((1 + XIRR_VAL) - 1)
Else
    SUMMARY_VECTOR(1, 4) = "N/A"
End If
SUMMARY_VECTOR(1, 5) = FACTOR_VAL * (DATA_MATRIX(NROWS, 7) - DATA_MATRIX(1, 6) - NET_CONTRIBUTIONS_VAL) / (DATA_MATRIX(1, 6) + CONTRIBUTIONS_VAL)

'-----------------------------------------------------------------------------------
Select Case OUTPUT
'-----------------------------------------------------------------------------------
Case 0 'Perfect
'-----------------------------------------------------------------------------------
    Erase h: Erase SUMMARY_VECTOR
    Erase DATE2_VECTOR: Erase DATA2_VECTOR
    ASSET_MTMDR_FUNC = DATA_MATRIX
'-----------------------------------------------------------------------------------
Case Else
'-----------------------------------------------------------------------------------
    For k = 6 To 10: SUMMARY_VECTOR(1, k) = 1: Next k
    
    ReDim DATA_VECTOR(0 To l, 1 To 16)
    DATA_VECTOR(0, 1) = "BEGIN DATE"
    DATA_VECTOR(0, 2) = "END DATE"
    DATA_VECTOR(0, 3) = "BEGIN IDX"
    DATA_VECTOR(0, 4) = "END IDX"
    DATA_VECTOR(0, 5) = "DAYS IN MONTH"
    DATA_VECTOR(0, 6) = "BMV"
    DATA_VECTOR(0, 7) = "EMV"
    DATA_VECTOR(0, 8) = "CONTRIBUTIONS"
    DATA_VECTOR(0, 9) = "WITHDRAWALS"
    DATA_VECTOR(0, 10) = "NET CONTRIBUTIONS"
    DATA_VECTOR(0, 11) = "WEIGHTED NET CONTRIBUTIONS"
    DATA_VECTOR(0, 12) = "DIETZ"
    DATA_VECTOR(0, 13) = "MODIFIED DIETZ"
    DATA_VECTOR(0, 14) = "TRUE TWR"
    DATA_VECTOR(0, 15) = "XIRR"
    DATA_VECTOR(0, 16) = "ROI"
    For k = 1 To l
        i = h(1, k)
        DATA_VECTOR(k, 1) = DATA_MATRIX(i, 1)
        DATA_VECTOR(k, 3) = i
        If k <> l Then
            j = h(1, k + 1) - 1
            DATA_VECTOR(k, 2) = DATA_MATRIX(j, 1)
            DATA_VECTOR(k, 4) = j
        Else
            j = NROWS
            DATA_VECTOR(k, 2) = DATA_MATRIX(j, 1)
            DATA_VECTOR(k, 4) = j
        End If
        DATA_VECTOR(k, 5) = DATA_VECTOR(k, 4) - DATA_VECTOR(k, 3) + 1
        DATA_VECTOR(k, 6) = DATA_MATRIX(i, 6)
        DATA_VECTOR(k, 7) = DATA_MATRIX(j, 7)
        DATA_VECTOR(k, 8) = 0
        DATA_VECTOR(k, 9) = 0
        DATA_VECTOR(k, 11) = 0
        DATA_VECTOR(k, 14) = 1
        
        n = j - i + 1
        ReDim DATE2_VECTOR(1 To n, 1 To 1)
        ReDim DATA2_VECTOR(1 To n, 1 To 1)
        
        n = 1
        For m = i To j
            If DATA_MATRIX(m, 5) > 0 Then
                DATA_VECTOR(k, 8) = DATA_VECTOR(k, 8) + DATA_MATRIX(m, 5)
            ElseIf DATA_MATRIX(m, 5) < 0 Then
                DATA_VECTOR(k, 9) = DATA_VECTOR(k, 9) + DATA_MATRIX(m, 5)
            End If
            DATA_VECTOR(k, 11) = DATA_VECTOR(k, 11) + DATA_MATRIX(m, 9)
            DATA_VECTOR(k, 14) = DATA_VECTOR(k, 14) * DATA_MATRIX(m, 12)
            DATE2_VECTOR(n, 1) = DATA_MATRIX(m, 1)
            DATA2_VECTOR(n, 1) = DATA_MATRIX(m, 15)
            n = n + 1
        Next m
        DATA_VECTOR(k, 14) = (DATA_VECTOR(k, 14) - 1) * FACTOR_VAL

        XIRR_VAL = IRR_FUNC(DATA2_VECTOR, DATE2_VECTOR, FACTOR_VAL, 1000, 10 ^ -10)
        If XIRR_VAL <> PUB_EPSILON Then
            DATA_VECTOR(k, 15) = XIRR_VAL
            DATA_VECTOR(k, 15) = FACTOR_VAL * ((1 + DATA_VECTOR(k, 15)) ^ ((DATA_VECTOR(k, 4) - DATA_VECTOR(k, 3)) / (365)) - 1)
        Else
            DATA_VECTOR(k, 15) = "N/A"
        End If

        DATA_VECTOR(k, 10) = DATA_VECTOR(k, 8) + DATA_VECTOR(k, 9)
        If (DATA_VECTOR(k, 6) + 0.5 * DATA_VECTOR(k, 10)) <> 0 Then
            DATA_VECTOR(k, 12) = FACTOR_VAL * (DATA_VECTOR(k, 7) - DATA_VECTOR(k, 6) - _
                                 DATA_VECTOR(k, 10)) / (DATA_VECTOR(k, 6) + _
                                 0.5 * DATA_VECTOR(k, 10))
        End If
        If (DATA_VECTOR(k, 6) + DATA_VECTOR(k, 11)) <> 0 Then
            DATA_VECTOR(k, 13) = FACTOR_VAL * (DATA_VECTOR(k, 7) - DATA_VECTOR(k, 6) - _
                                 DATA_VECTOR(k, 10)) / (DATA_VECTOR(k, 6) + _
                                 DATA_VECTOR(k, 11))
        End If
        
        If (DATA_VECTOR(k, 6) + DATA_VECTOR(k, 8)) <> 0 Then
            DATA_VECTOR(k, 16) = FACTOR_VAL * ((DATA_VECTOR(k, 7) - DATA_VECTOR(k, 9)) - _
                                (DATA_VECTOR(k, 6) + DATA_VECTOR(k, 8))) / _
                                (DATA_VECTOR(k, 6) + DATA_VECTOR(k, 8))
        End If
        
        SUMMARY_VECTOR(1, 6) = SUMMARY_VECTOR(1, 6) * (1 + DATA_VECTOR(k, 12) / FACTOR_VAL)
        SUMMARY_VECTOR(1, 7) = SUMMARY_VECTOR(1, 7) * (1 + DATA_VECTOR(k, 13) / FACTOR_VAL)
        SUMMARY_VECTOR(1, 8) = SUMMARY_VECTOR(1, 8) * (1 + DATA_VECTOR(k, 14) / FACTOR_VAL)
        If IsNumeric(DATA_VECTOR(k, 15)) Then
            SUMMARY_VECTOR(1, 9) = SUMMARY_VECTOR(1, 9) * (1 + DATA_VECTOR(k, 15) / FACTOR_VAL)
        End If
        SUMMARY_VECTOR(1, 10) = SUMMARY_VECTOR(1, 10) * (1 + DATA_VECTOR(k, 16) / FACTOR_VAL)
    
    Next k

    If OUTPUT = 1 Then
        Erase h
        Erase DATA_MATRIX: Erase SUMMARY_VECTOR
        Erase DATE2_VECTOR: Erase DATA2_VECTOR
        ASSET_MTMDR_FUNC = DATA_VECTOR
        Exit Function
    End If

    SUMMARY_VECTOR(1, 6) = FACTOR_VAL * (SUMMARY_VECTOR(1, 6) - 1)
    SUMMARY_VECTOR(1, 7) = FACTOR_VAL * (SUMMARY_VECTOR(1, 7) - 1)
    SUMMARY_VECTOR(1, 8) = FACTOR_VAL * (SUMMARY_VECTOR(1, 8) - 1)
    SUMMARY_VECTOR(1, 9) = FACTOR_VAL * (SUMMARY_VECTOR(1, 9) - 1)
    SUMMARY_VECTOR(1, 10) = FACTOR_VAL * (SUMMARY_VECTOR(1, 10) - 1)
    
    If OUTPUT = 2 Then
        ASSET_MTMDR_FUNC = SUMMARY_VECTOR
    Else
        Erase h
        Erase DATE2_VECTOR: Erase DATA2_VECTOR
        ASSET_MTMDR_FUNC = Array(DATA_MATRIX, DATA_VECTOR, SUMMARY_VECTOR)
    End If
'-----------------------------------------------------------------------------------
End Select
'-----------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
ASSET_MTMDR_FUNC = Err.number
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


