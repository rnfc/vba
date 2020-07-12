Attribute VB_Name = "FINAN_FUNDAM_CAR_LEASING_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : CAR_LEASING_FUNC
'DESCRIPTION   : CAR LEASING MODEL
'LIBRARY       : FUNDAMENTAL
'GROUP         : CAR_LEASING
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/09/2010
'REFERENCE     : http://www.financialwebring.org/gummy-stuff/Car-Leasing.htm
'************************************************************************************
'************************************************************************************

Function CAR_LEASING_FUNC( _
ByVal MSRP_RNG As Variant, _
ByVal CAPITALIZED_COST_RNG As Variant, _
ByVal DOWN_PAYMENT_RNG As Variant, _
ByVal RESIDUAL_FACTOR_RNG As Variant, _
ByVal INTEREST_RATE_RNG As Variant, _
Optional ByVal MONTHS_RNG As Variant = 36, _
Optional ByVal OUTPUT As Integer = 1)

'-------------------------------------------------------------------------------------
'Manufacturer 's Suggested Retail Price (MSRP) =  $30,000
'Negotiated Price (Capitalized Cost) A = $27,000
'Down payment or Trade-in (Capitalized Cost Reduction) D  =  $4,000
'Worth, after Depreciation (Residual Factor) f = 50%
'Term (months) M  =  36
'Annual Interest Rate I =    8.00%
'-------------------------------------------------------------------------------------
'RESIDUAL VALUE = Residual Factor x MSRP    (value of car at end of lease)
'MONTHLY DEPRECIATION FEE = (COST after REDUCTION - Residual Value) / LEASE TERM
'MONTHLY LEASE FEE = (COST after REDUCTION + Residual Value) x Interest Rate /24
'"MONEY FACTOR" = (Interest Rate)/24   and  ANNUAL INTEREST RATE is expressed as a percentage
'TOTAL MONTHLY LEASE PAYMENT = Monthly Depreciation Fee + Monthly Leasing Fee
'-------------------------------------------------------------------------------------
'[Do Sensitivity Analysis Based on Interest Rate]
'[Remember to add a percentage for taxes in the monthly payment, etc]
'-------------------------------------------------------------------------------------

Dim i As Long
Dim NROWS As Long
Dim TEMP_STR As String
Dim PAYMENT_VAL As Double
Dim MSRP_VECTOR As Variant
Dim CAPITALIZED_COST_VECTOR As Variant
Dim DOWN_PAYMENT_VECTOR As Variant
Dim RESIDUAL_FACTOR_VECTOR As Variant
Dim MONTHS_VECTOR As Variant
Dim INTEREST_RATE_VECTOR As Variant

Dim TEMP_MATRIX As Variant
Dim TEMP_VECTOR() As String

'Debug.Print CAR_LEASING_FUNC(Range("K2:S2"), Range("K3:S3"), Range("K4:S4"), Range("K5:S5"), Range("K7:S7"), Range("K8:S8"))(1, 1)

On Error GoTo ERROR_LABEL

If IsArray(MSRP_RNG) = True Then
    MSRP_VECTOR = MSRP_RNG
    If UBound(MSRP_VECTOR, 1) = 1 Then
        MSRP_VECTOR = MATRIX_TRANSPOSE_FUNC(MSRP_VECTOR)
    End If
Else
    ReDim MSRP_VECTOR(1 To 1, 1 To 1)
    MSRP_VECTOR(1, 1) = MSRP_RNG
End If
NROWS = UBound(MSRP_VECTOR, 1)

If IsArray(CAPITALIZED_COST_RNG) = True Then
    CAPITALIZED_COST_VECTOR = CAPITALIZED_COST_RNG
    If UBound(CAPITALIZED_COST_VECTOR, 1) = 1 Then
        CAPITALIZED_COST_VECTOR = MATRIX_TRANSPOSE_FUNC(CAPITALIZED_COST_VECTOR)
    End If
Else
    ReDim CAPITALIZED_COST_VECTOR(1 To 1, 1 To 1)
    CAPITALIZED_COST_VECTOR(1, 1) = CAPITALIZED_COST_RNG
End If
If NROWS <> UBound(CAPITALIZED_COST_VECTOR, 1) Then: GoTo ERROR_LABEL

If IsArray(DOWN_PAYMENT_RNG) = True Then
    DOWN_PAYMENT_VECTOR = DOWN_PAYMENT_RNG
    If UBound(DOWN_PAYMENT_VECTOR, 1) = 1 Then
        DOWN_PAYMENT_VECTOR = MATRIX_TRANSPOSE_FUNC(DOWN_PAYMENT_VECTOR)
    End If
Else
    ReDim DOWN_PAYMENT_VECTOR(1 To 1, 1 To 1)
    DOWN_PAYMENT_VECTOR(1, 1) = DOWN_PAYMENT_RNG
End If
If NROWS <> UBound(DOWN_PAYMENT_VECTOR, 1) Then: GoTo ERROR_LABEL

If IsArray(RESIDUAL_FACTOR_RNG) = True Then
    RESIDUAL_FACTOR_VECTOR = RESIDUAL_FACTOR_RNG
    If UBound(RESIDUAL_FACTOR_VECTOR, 1) = 1 Then
        RESIDUAL_FACTOR_VECTOR = MATRIX_TRANSPOSE_FUNC(RESIDUAL_FACTOR_VECTOR)
    End If
Else
    ReDim RESIDUAL_FACTOR_VECTOR(1 To 1, 1 To 1)
    RESIDUAL_FACTOR_VECTOR(1, 1) = RESIDUAL_FACTOR_RNG
End If
If NROWS <> UBound(RESIDUAL_FACTOR_VECTOR, 1) Then: GoTo ERROR_LABEL

If IsArray(INTEREST_RATE_RNG) = True Then
    INTEREST_RATE_VECTOR = INTEREST_RATE_RNG
    If UBound(INTEREST_RATE_VECTOR, 1) = 1 Then
        INTEREST_RATE_VECTOR = MATRIX_TRANSPOSE_FUNC(INTEREST_RATE_VECTOR)
    End If
Else
    ReDim INTEREST_RATE_VECTOR(1 To 1, 1 To 1)
    INTEREST_RATE_VECTOR(1, 1) = INTEREST_RATE_RNG
End If
If NROWS <> UBound(INTEREST_RATE_VECTOR, 1) Then: GoTo ERROR_LABEL

If IsArray(MONTHS_RNG) = True Then
    MONTHS_VECTOR = MONTHS_RNG
    If UBound(MONTHS_VECTOR, 1) = 1 Then
        MONTHS_VECTOR = MATRIX_TRANSPOSE_FUNC(MONTHS_VECTOR)
    End If
Else
    ReDim MONTHS_VECTOR(1 To NROWS, 1 To 1)
    For i = 1 To NROWS
        MONTHS_VECTOR(i, 1) = MONTHS_RNG
    Next i
End If
If NROWS <> UBound(MONTHS_VECTOR, 1) Then: GoTo ERROR_LABEL


'----------------------------------------------------------------------------
'For the Sensitivity Analysis uses Monthly Lease Payments in the y-axis and
'Interest Rate in the x-axis.
'----------------------------------------------------------------------------

'----------------------------------------------------------------------------
Select Case OUTPUT
'----------------------------------------------------------------------------
Case 0
'----------------------------------------------------------------------------
    ReDim TEMP_MATRIX(0 To NROWS, 1 To 16)
    TEMP_MATRIX(0, 1) = "MSRP"
    TEMP_MATRIX(0, 2) = "CAPITALIZED COST"
    TEMP_MATRIX(0, 3) = "DOWN PAYMENT OR TRADE-IN"
    TEMP_MATRIX(0, 4) = "RESIDUAL FACTOR"
    TEMP_MATRIX(0, 5) = "INTEREST RATE"
    TEMP_MATRIX(0, 6) = "MONTHS"
    TEMP_MATRIX(0, 7) = "RESIDUAL VALUE"
    TEMP_MATRIX(0, 8) = "DEPRECIATION"
    TEMP_MATRIX(0, 9) = "COST AFTER REDUCTION"
    TEMP_MATRIX(0, 10) = "MONTHLY DEPRECIATION FEE"
    '--------------------------------------------------------------------------------------------------
    TEMP_MATRIX(0, 11) = "MONEY FACTOR"
    TEMP_MATRIX(0, 12) = "MONTHLY LEASING FEE"
    TEMP_MATRIX(0, 13) = "MONTHLY LEASE PAYMENTS (USING DEALER'S FORMULA)" 'Before Taxes
    TEMP_MATRIX(0, 14) = "LEASING LOAN"
    '--------------------------------------------------------------------------------------------------
    TEMP_MATRIX(0, 15) = "MONTHLY LEASE PAYMENTS (MORE ACCURATE FORMULA)" 'Before Taxes
    TEMP_MATRIX(0, 16) = "MONTHLY LEASE PAYMENTS (USING BANK'S FORMULA)" 'Before Taxes
    '--------------------------------------------------------------------------------------------------
    
    For i = 1 To NROWS
        If INTEREST_RATE_VECTOR(i, 1) = 0 Then: INTEREST_RATE_VECTOR(i, 1) = 10 ^ -10
        TEMP_MATRIX(i, 1) = MSRP_VECTOR(i, 1)
        TEMP_MATRIX(i, 2) = CAPITALIZED_COST_VECTOR(i, 1)
        TEMP_MATRIX(i, 3) = DOWN_PAYMENT_VECTOR(i, 1)
        TEMP_MATRIX(i, 4) = RESIDUAL_FACTOR_VECTOR(i, 1)
        TEMP_MATRIX(i, 5) = INTEREST_RATE_VECTOR(i, 1)
        TEMP_MATRIX(i, 6) = MONTHS_VECTOR(i, 1)
        TEMP_MATRIX(i, 7) = TEMP_MATRIX(i, 1) * TEMP_MATRIX(i, 4)
        TEMP_MATRIX(i, 8) = TEMP_MATRIX(i, 1) - TEMP_MATRIX(i, 7)
        TEMP_MATRIX(i, 9) = TEMP_MATRIX(i, 2) - TEMP_MATRIX(i, 3)
        TEMP_MATRIX(i, 10) = (TEMP_MATRIX(i, 9) - TEMP_MATRIX(i, 7)) / TEMP_MATRIX(i, 6)
        TEMP_MATRIX(i, 11) = TEMP_MATRIX(i, 5) / 24
        TEMP_MATRIX(i, 12) = (TEMP_MATRIX(i, 7) + TEMP_MATRIX(i, 9)) * TEMP_MATRIX(i, 11)

        TEMP_MATRIX(i, 13) = (TEMP_MATRIX(i, 2) - TEMP_MATRIX(i, 3) - TEMP_MATRIX(i, 7)) / TEMP_MATRIX(i, 6) + (TEMP_MATRIX(i, 2) - TEMP_MATRIX(i, 3) + TEMP_MATRIX(i, 7)) * TEMP_MATRIX(i, 5) / 24
        'Debug.Print TEMP_MATRIX(i, 13) = (TEMP_MATRIX(i, 12) + TEMP_MATRIX(i, 10))
        TEMP_MATRIX(i, 14) = 12 * TEMP_MATRIX(i, 13) * (1 - (1 + INTEREST_RATE_VECTOR(i, 1) / 12) ^ (-MONTHS_VECTOR(i, 1))) / INTEREST_RATE_VECTOR(i, 1)
        
        If ((1 + TEMP_MATRIX(i, 5) / 12) ^ TEMP_MATRIX(i, 6) - 1) <> 0 Then
            TEMP_MATRIX(i, 15) = (TEMP_MATRIX(i, 5) / 12) * ((TEMP_MATRIX(i, 2) - TEMP_MATRIX(i, 3)) * (1 + TEMP_MATRIX(i, 5)) ^ (TEMP_MATRIX(i, 6) / 12) - TEMP_MATRIX(i, 7)) / ((1 + TEMP_MATRIX(i, 5) / 12) ^ TEMP_MATRIX(i, 6) - 1)
        Else
            TEMP_MATRIX(i, 15) = CVErr(xlErrNA)
        End If
        If ((1 + TEMP_MATRIX(i, 5)) ^ (TEMP_MATRIX(i, 6) / 12) - 1) <> 0 Then
            TEMP_MATRIX(i, 16) = ((TEMP_MATRIX(i, 2) - TEMP_MATRIX(i, 3)) * (1 + TEMP_MATRIX(i, 5)) ^ (TEMP_MATRIX(i, 6) / 12) - TEMP_MATRIX(i, 7)) * ((1 + TEMP_MATRIX(i, 5)) ^ (1 / 12) - 1) / ((1 + TEMP_MATRIX(i, 5)) ^ (TEMP_MATRIX(i, 6) / 12) - 1)
        Else
            TEMP_MATRIX(i, 16) = CVErr(xlErrNA)
        End If
    Next i
    CAR_LEASING_FUNC = TEMP_MATRIX
'----------------------------------------------------------------------------
Case Else
'----------------------------------------------------------------------------
    ReDim TEMP_VECTOR(1 To NROWS, 1 To 1)
    For i = 1 To NROWS
        TEMP_STR = "": GoSub REPORT_LINE
        TEMP_VECTOR(i, 1) = TEMP_STR
    Next i
    CAR_LEASING_FUNC = TEMP_VECTOR
'----------------------------------------------------------------------------
End Select
'----------------------------------------------------------------------------

Exit Function
'-------------------------------------------------------------------------------------------------------------------------------------
REPORT_LINE:
'-------------------------------------------------------------------------------------------------------------------------------------
    If INTEREST_RATE_VECTOR(i, 1) = 0 Then: INTEREST_RATE_VECTOR(i, 1) = 10 ^ -10
    TEMP_STR = "Compare to borrowing the money: " & Format(CAPITALIZED_COST_VECTOR(i, 1) - DOWN_PAYMENT_VECTOR(i, 1), "0,#.00") & " at " & Format(100 * INTEREST_RATE_VECTOR(i, 1), "0.0") & "% / 12 per month so that the balance is " & Format(MSRP_VECTOR(i, 1) * RESIDUAL_FACTOR_VECTOR(i, 1), "0,#.00") & " after " & Format(MONTHS_VECTOR(i, 1), "0") & " months. "
    TEMP_STR = TEMP_STR & "Amount borrowed to buy the car (after down payment): " & Format(CAPITALIZED_COST_VECTOR(i, 1) - DOWN_PAYMENT_VECTOR(i, 1), "0,#.00") & ". "
    TEMP_STR = TEMP_STR & "Residual value of car after " & Format(MONTHS_VECTOR(i, 1), "0") & " months: " & Format(MSRP_VECTOR(i, 1) * RESIDUAL_FACTOR_VECTOR(i, 1), "0,#.00") & ". "
    TEMP_STR = TEMP_STR & "Monthly Bank Payments " & Format(-(CAPITALIZED_COST_VECTOR(i, 1) - DOWN_PAYMENT_VECTOR(i, 1) - MSRP_VECTOR(i, 1) * RESIDUAL_FACTOR_VECTOR(i, 1) * (1 + INTEREST_RATE_VECTOR(i, 1) / 12) ^ (-MONTHS_VECTOR(i, 1))) / ((1 - (1 / ((1 + (INTEREST_RATE_VECTOR(i, 1) / 12)) ^ MONTHS_VECTOR(i, 1)))) / (INTEREST_RATE_VECTOR(i, 1) / 12)), "0,#.00") & " so that the balance due is " & Format(MSRP_VECTOR(i, 1) * RESIDUAL_FACTOR_VECTOR(i, 1), "0,#.00") & " after " & Format(MONTHS_VECTOR(i, 1), "0") & " months. "
    TEMP_STR = TEMP_STR & "Scenario (1) The dealer sells the car and gets " & Format(CAPITALIZED_COST_VECTOR(i, 1), "0,#.00") & ". After " & Format(MONTHS_VECTOR(i, 1) / 12, "0") & " years at " & Format(INTEREST_RATE_VECTOR(i, 1), "0.00%") & " per year, this is worth " & Format(CAPITALIZED_COST_VECTOR(i, 1) * (1 + INTEREST_RATE_VECTOR(i, 1) / 12) ^ MONTHS_VECTOR(i, 1), "0,#.00") & " to the dealer. "
    PAYMENT_VAL = ((CAPITALIZED_COST_VECTOR(i, 1) - DOWN_PAYMENT_VECTOR(i, 1)) - (MSRP_VECTOR(i, 1) * RESIDUAL_FACTOR_VECTOR(i, 1))) / MONTHS_VECTOR(i, 1) + ((MSRP_VECTOR(i, 1) * RESIDUAL_FACTOR_VECTOR(i, 1)) + (CAPITALIZED_COST_VECTOR(i, 1) - DOWN_PAYMENT_VECTOR(i, 1))) * (INTEREST_RATE_VECTOR(i, 1) / 24)
    TEMP_STR = TEMP_STR & "Scenario (2) They lease the car, receiving " & Format(PAYMENT_VAL, "0,#.00") & " per month. Suppose they invest these payments at " & Format(INTEREST_RATE_VECTOR(i, 1), "0.00%") & " per year. "
    TEMP_STR = TEMP_STR & Format(MONTHS_VECTOR(i, 1), "0") & " months of these investments are worth " & Format(12 * PAYMENT_VAL * ((1 + INTEREST_RATE_VECTOR(i, 1) / 12) ^ MONTHS_VECTOR(i, 1) - 1) / INTEREST_RATE_VECTOR(i, 1), "0,#.00") & ". The " & Format(DOWN_PAYMENT_VECTOR(i, 1), "0,#.00") & " down payment is worth " & Format(DOWN_PAYMENT_VECTOR(i, 1) * (1 + INTEREST_RATE_VECTOR(i, 1) / 12) ^ MONTHS_VECTOR(i, 1), "0,#.00") & " after " & Format(MONTHS_VECTOR(i, 1) / 12, "0") & " years (at " & Format(INTEREST_RATE_VECTOR(i, 1), "0.00%") & " per year). "
    TEMP_STR = TEMP_STR & "Adding: " & Format((RESIDUAL_FACTOR_VECTOR(i, 1) * MSRP_VECTOR(i, 1)), "0,#.00") & " (residual value) + " & Format(12 * PAYMENT_VAL * ((1 + INTEREST_RATE_VECTOR(i, 1) / 12) ^ MONTHS_VECTOR(i, 1) - 1) / INTEREST_RATE_VECTOR(i, 1), "0,#.00") & " (monthly investments value) + " & Format(DOWN_PAYMENT_VECTOR(i, 1) * (1 + INTEREST_RATE_VECTOR(i, 1) / 12) ^ MONTHS_VECTOR(i, 1), "0,#.00") & " (down payment value) = " & Format(12 * PAYMENT_VAL * ((1 + INTEREST_RATE_VECTOR(i, 1) / 12) ^ MONTHS_VECTOR(i, 1) - 1) / INTEREST_RATE_VECTOR(i, 1) + (RESIDUAL_FACTOR_VECTOR(i, 1) * MSRP_VECTOR(i, 1)) + DOWN_PAYMENT_VECTOR(i, 1) * (1 + INTEREST_RATE_VECTOR(i, 1) / 12) ^ MONTHS_VECTOR(i, 1), "0,#.00") & ". "
    TEMP_STR = TEMP_STR & "These two numbers, " & Format(CAPITALIZED_COST_VECTOR(i, 1) * (1 + INTEREST_RATE_VECTOR(i, 1) / 12) ^ MONTHS_VECTOR(i, 1), "0,#.00") & " and " & Format(12 * PAYMENT_VAL * ((1 + INTEREST_RATE_VECTOR(i, 1) / 12) ^ MONTHS_VECTOR(i, 1) - 1) / INTEREST_RATE_VECTOR(i, 1) + MSRP_VECTOR(i, 1) * RESIDUAL_FACTOR_VECTOR(i, 1) + DOWN_PAYMENT_VECTOR(i, 1) * (1 + INTEREST_RATE_VECTOR(i, 1) / 12) ^ MONTHS_VECTOR(i, 1), "0,#.00") & ", should be about the same."
'-------------------------------------------------------------------------------------------------------------------------------------
Return
'-------------------------------------------------------------------------------------------------------------------------------------
ERROR_LABEL:
CAR_LEASING_FUNC = Err.number
End Function
