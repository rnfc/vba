Attribute VB_Name = "FINAN_FI_BOND_MORTGAGES_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : MORTGAGE_PAYMENT_FUNC

'LIBRARY       : MORTGAGES
'GROUP         : FIXED-INCOME
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'************************************************************************************
'************************************************************************************

Function MORTGAGE_PAYMENT_FUNC(ByVal AMOUNT_VAL As Double, _
ByVal RATE_VAL As Double, _
ByVal YEARS_VAL As Double, _
Optional ByVal DATE_VAL As Double = 0, _
Optional ByVal OUTPUT As Integer = 30)

Dim h As Long 'Days
Dim i As Long
Dim j As Long
Dim k As Long 'Sum Days
Dim l As Long 'Days left

Dim TEMP_SUM As Double 'Sum Interests
Dim TEMP1_VAL As Double 'Temp Payments

Dim RATE1_VAL As Double 'U.S. Daily Rate
Dim RATE2_VAL As Double 'Canada Daily Rate
Dim RATE3_VAL As Double 'Nico Daily Rate

Dim PAYMENT1_VAL As Double 'U.S.
Dim PAYMENT2_VAL As Double 'Canada
Dim PAYMENT3_VAL As Double 'Nico

Dim BALANCE_VAL As Double 'Balance Due
Dim TEMP2_VAL As Double 'Payments Sum

On Error GoTo ERROR_LABEL

RATE1_VAL = RATE_VAL / 12
PAYMENT1_VAL = AMOUNT_VAL / ((1 - (1 / ((1 + RATE1_VAL) ^ (YEARS_VAL * 12)))) / RATE1_VAL)

RATE2_VAL = (1 + RATE_VAL / 2) ^ (1 / 6) - 1
PAYMENT2_VAL = AMOUNT_VAL / ((1 - (1 / ((1 + RATE2_VAL) ^ (YEARS_VAL * 12)))) / RATE2_VAL)

RATE3_VAL = (1 + RATE_VAL) ^ (1 / 365) - 1
GoSub FACTOR1_LINE
PAYMENT3_VAL = AMOUNT_VAL * RATE_VAL / TEMP_SUM / (1 - 1 / ((1 + RATE_VAL)) ^ YEARS_VAL)

'--------------------------------------------------------------------------------------------------------
Select Case OUTPUT
'--------------------------------------------------------------------------------------------------------
Case 0
'--------------------------------------------------------------------------------------------------------
    MORTGAGE_PAYMENT_FUNC = Array(PAYMENT1_VAL, PAYMENT2_VAL, PAYMENT3_VAL)
'--------------------------------------------------------------------------------------------------------
Case Else
'--------------------------------------------------------------------------------------------------------
    GoSub FACTOR2_LINE
    j = OUTPUT
    For i = 1 To j
        If i > 1 Then
            TEMP2_VAL = TEMP2_VAL * (1 + RATE_VAL) + TEMP1_VAL
        Else
            TEMP2_VAL = TEMP1_VAL
        End If
        BALANCE_VAL = AMOUNT_VAL * (1 + RATE_VAL) ^ i - TEMP2_VAL
        If BALANCE_VAL < 0 Then: BALANCE_VAL = 0
    Next i
    MORTGAGE_PAYMENT_FUNC = Array(j, TEMP2_VAL, BALANCE_VAL)
'--------------------------------------------------------------------------------------------------------
End Select
'--------------------------------------------------------------------------------------------------------

Exit Function
'--------------------------------------------------------------------------------------------------------
FACTOR1_LINE:
'--------------------------------------------------------------------------------------------------------
k = 0
If DATE_VAL = 0 Then: DATE_VAL = Now
TEMP_SUM = 0
For i = 1 To 12
    GoSub DATE_LINE
    k = k + h
    l = 365 - k
    TEMP_SUM = TEMP_SUM + (1 + RATE3_VAL) ^ l
Next i
'--------------------------------------------------------------------------------------------------------
FACTOR2_LINE:
'--------------------------------------------------------------------------------------------------------
If DATE_VAL = 0 Then: DATE_VAL = Now
TEMP1_VAL = PAYMENT3_VAL
i = 1
GoSub DATE_LINE
j = h
For i = 2 To 12
    GoSub DATE_LINE
    TEMP1_VAL = TEMP1_VAL * (1 + RATE3_VAL) ^ j + PAYMENT3_VAL
    j = h
Next i
'--------------------------------------------------------------------------------------------------------
Return
'--------------------------------------------------------------------------------------------------------
DATE_LINE:
'--------------------------------------------------------------------------------------------------------=1
h = DateDiff("d", DateSerial(Year(DATE_VAL), Month(DATE_VAL) + i - 1, Day(DATE_VAL)), DateSerial(Year(DATE_VAL), Month(DATE_VAL) + i, Day(DATE_VAL)))
'--------------------------------------------------------------------------------------------------------
Return
'--------------------------------------------------------------------------------------------------------
ERROR_LABEL:
MORTGAGE_PAYMENT_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : FIXED_RATE_MORTGAGE_FUNC
'DESCRIPTION   : Fixed-rate Mortgage (FRM) Function
'LIBRARY       : MORTGAGES
'GROUP         : FIXED-INCOME
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'************************************************************************************
'************************************************************************************

Function FIXED_RATE_MORTGAGE_FUNC(ByVal LOAN_AMOUNT As Double, _
ByVal FIRST_PAYMENT As Date, _
ByVal TENOR As Double, _
ByVal RATE As Double, _
Optional ByVal FREQUENCY As Integer = 12, _
Optional ByVal OUTPUT As Integer = 0)

' RATE MUST BE ANNUALIZED

Dim i As Long
Dim j As Long 'periods

Dim TEMP_SUM As Double
Dim PAYMENT_VAL As Double

Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

j = Int(FREQUENCY * TENOR)
PAYMENT_VAL = ANNUITY_FUNC(0, LOAN_AMOUNT, RATE / FREQUENCY, j)

ReDim TEMP_MATRIX(0 To j, 1 To 8)

TEMP_MATRIX(0, 1) = "PERIODS"
TEMP_MATRIX(0, 2) = "MATURITY"
TEMP_MATRIX(0, 3) = "INTEREST"
TEMP_MATRIX(0, 4) = "PRINCIPAL"
TEMP_MATRIX(0, 5) = "REMAINING"
TEMP_MATRIX(0, 6) = "CUMULATIVE"
TEMP_MATRIX(0, 7) = "DISCOUNT"
TEMP_MATRIX(0, 8) = "INTRINSIC"

TEMP_MATRIX(1, 1) = 1
TEMP_MATRIX(1, 2) = EDATE_FUNC(FIRST_PAYMENT, TEMP_MATRIX(1, 1) * 12 / FREQUENCY)
TEMP_MATRIX(1, 3) = LOAN_AMOUNT * RATE / FREQUENCY
TEMP_MATRIX(1, 4) = PAYMENT_VAL - TEMP_MATRIX(1, 3)
TEMP_MATRIX(1, 5) = LOAN_AMOUNT - TEMP_MATRIX(1, 4)
TEMP_MATRIX(1, 6) = PAYMENT_VAL * TEMP_MATRIX(1, 1)
TEMP_MATRIX(1, 7) = 1 / (1 + (RATE / FREQUENCY)) ^ TEMP_MATRIX(1, 1)
TEMP_MATRIX(1, 8) = TEMP_MATRIX(1, 7) * TEMP_MATRIX(1, 6)

TEMP_SUM = TEMP_MATRIX(1, 8)

For i = 2 To j
    
    TEMP_MATRIX(i, 1) = i
    TEMP_MATRIX(i, 2) = EDATE_FUNC(FIRST_PAYMENT, TEMP_MATRIX(i, 1) * 12 / FREQUENCY)
    TEMP_MATRIX(i, 3) = TEMP_MATRIX(i - 1, 5) * RATE / FREQUENCY
    TEMP_MATRIX(i, 4) = PAYMENT_VAL - TEMP_MATRIX(i, 3)
    TEMP_MATRIX(i, 5) = TEMP_MATRIX(i - 1, 5) - TEMP_MATRIX(i, 4)
    TEMP_MATRIX(i, 6) = PAYMENT_VAL * TEMP_MATRIX(i, 1)
    TEMP_MATRIX(i, 7) = 1 / (1 + (RATE / FREQUENCY)) ^ TEMP_MATRIX(i, 1)
    TEMP_MATRIX(i, 8) = TEMP_MATRIX(i, 7) * TEMP_MATRIX(i, 6)
    
    TEMP_SUM = TEMP_SUM + TEMP_MATRIX(i, 8)
Next i

'----------------------------------------------------------------------------
Select Case OUTPUT
'----------------------------------------------------------------------------
Case 0
'----------------------------------------------------------------------------
    FIXED_RATE_MORTGAGE_FUNC = TEMP_MATRIX
'----------------------------------------------------------------------------
Case 1
'----------------------------------------------------------------------------
    FIXED_RATE_MORTGAGE_FUNC = TEMP_SUM 'Sum of time-weighted PV pmts
'---------------------Critical: Assume no pre-payment------------------------
'--> Prepayment is an option given to the borrower to put the
'--> loan at par to the lender.
Case 2
'----------------------------------------------------------------------------
    FIXED_RATE_MORTGAGE_FUNC = TEMP_SUM / LOAN_AMOUNT / FREQUENCY 'Macaulay duration
'----------------------------------------------------------------------------
Case Else
'----------------------------------------------------------------------------
    FIXED_RATE_MORTGAGE_FUNC = (TEMP_SUM / LOAN_AMOUNT / FREQUENCY) / ((1 + RATE / FREQUENCY) ^ FREQUENCY)
    'Macaulay duration. Prepayments shorten the expected duration and greatly complicate the valuation of MBS.
'----------------------------------------------------------------------------
End Select
'----------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
FIXED_RATE_MORTGAGE_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : ADJUSTABLE_RATE_MORTGAGE_FUNC
'DESCRIPTION   : Adjustable Rate Mortgage (ARM) Function
'LIBRARY       : MORTGAGES
'GROUP         : FIXED-INCOME
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'************************************************************************************
'************************************************************************************

Function ADJUSTABLE_RATE_MORTGAGE_FUNC(ByVal LOAN_AMOUNT As Double, _
ByVal TENOR As Double, _
ByVal START_RATE As Double, _
ByVal FIRST_PAYMENT As Date, _
ByVal FIXED_TENOR As Double, _
ByVal ADJUSTED_TENOR As Double, _
ByVal ADJUSTED_RATE As Double, _
ByVal CAP_RATE As Double, _
Optional ByRef ADDITIONAL_PAYMENTS_RNG As Variant, _
Optional ByVal FREQUENCY As Integer = 12, _
Optional ByVal OUTPUT As Integer = 0)

'-------------------------------CHARTS-------------------------------------

'CHART_A: x-axis(DATES) ; y-axis(Loan Balance and Cumul. values)
'CHART_B: x-axis(DATES); y-axis(Interest Rate evolution)

'--------------------------------------------------------------------------

'TENOR: Term in years --> Mortgage loans usually have 15 or 30-year terms.
'Auto loans are usually between 2 and 5 years.

'START_RATE: Starting interest rate --> The starting annual interest rate.
'For most popular ARMs, this rate remains fixed for a specified number of periods.

'FIRST_PAYMENT: First payment date --> Assumes that the first payment date
'is at the end of the first period.

'FIXED_TENOR: Rate remains fixed for x years --> In a 60-periods ARM,
'the initial interest rate remains fixed for the first 5 years (assuming a
'frequency of 12 periods per year). After that, the rate is subject to
'adjustments, depending upon market conditions.

'ADJUSTED_RATE: Expected adjustment in percentage (annualized)

'ADJUSTED_TENOR: Periods between adjustments. The adjustment period
'is the number of periods between each interest rate adjustment. The
'common adjustment period is per 12 months, meaning that the rate will
'be adjusted once a year at most.

'CAP_RATE: The maximum interest rate the mortgage allows.
'The rate will not be adjusted higher than the cap.

'ADDITIONAL_PAYMNETS: Additional Payments per period.

'---------------------------------------------------------------------------

Dim i As Long
Dim j As Long


Dim TEMP1_SUM As Double
Dim TEMP2_SUM As Double

Dim MAX_RATE As Double
Dim MAX_PAYMENT As Double

Dim PAYMENT_VAL As Double

Dim TEMP_VECTOR As Variant
Dim TEMP_MATRIX As Variant
Dim ADDITIONAL_VECTOR As Variant

On Error GoTo ERROR_LABEL

j = Int(TENOR * FREQUENCY)

If IsArray(ADDITIONAL_PAYMENTS_RNG) = True Then
    ADDITIONAL_VECTOR = ADDITIONAL_PAYMENTS_RNG
    If UBound(ADDITIONAL_VECTOR, 1) = 1 Then
        ADDITIONAL_VECTOR = MATRIX_TRANSPOSE_FUNC(ADDITIONAL_VECTOR)
    End If
End If

ReDim TEMP_MATRIX(0 To j + 1, 1 To 10)

TEMP_MATRIX(0, 1) = "PAYMENT #"
TEMP_MATRIX(0, 2) = "PAYMENT DATE"
TEMP_MATRIX(0, 3) = "INTEREST RATE"
TEMP_MATRIX(0, 4) = "PAYMENT DUE"
TEMP_MATRIX(0, 5) = "ADDITIONAL PAYMENT"
TEMP_MATRIX(0, 6) = "INTEREST"
TEMP_MATRIX(0, 7) = "PRINCIPAL"
TEMP_MATRIX(0, 8) = "BALANCE"
TEMP_MATRIX(0, 9) = "CUMUL. INTEREST"
TEMP_MATRIX(0, 10) = "CUMUL. PRINCIPAL"

TEMP_MATRIX(1, 1) = ""
TEMP_MATRIX(1, 2) = ""
TEMP_MATRIX(1, 3) = 0#
TEMP_MATRIX(1, 4) = 0#
TEMP_MATRIX(1, 5) = ""
TEMP_MATRIX(1, 6) = ""
TEMP_MATRIX(1, 7) = ""
TEMP_MATRIX(1, 8) = LOAN_AMOUNT
TEMP_MATRIX(1, 9) = ""
TEMP_MATRIX(1, 10) = ""

PAYMENT_VAL = ANNUITY_FUNC(0, LOAN_AMOUNT, START_RATE / FREQUENCY, TENOR * FREQUENCY)

TEMP1_SUM = 0
TEMP2_SUM = 0

MAX_RATE = TEMP_MATRIX(1, 3)
MAX_PAYMENT = TEMP_MATRIX(1, 4)

For i = 1 To j
    TEMP_MATRIX(i + 1, 1) = i
    
    TEMP_MATRIX(i + 1, 2) = EDATE_FUNC(FIRST_PAYMENT, (i - 1) * 12 / FREQUENCY)
    
    If TEMP_MATRIX(i + 1, 1) < (FIXED_TENOR * FREQUENCY) Then
        TEMP_MATRIX(i + 1, 3) = START_RATE
    Else
        TEMP_MATRIX(i + 1, 3) = MINIMUM_FUNC(CAP_RATE, START_RATE + ADJUSTED_RATE * (TEMP_MATRIX(i + 1, 1) - FIXED_TENOR * FREQUENCY) / (ADJUSTED_TENOR * FREQUENCY))
    End If

    MAX_RATE = MAXIMUM_FUNC(TEMP_MATRIX(i + 1, 3), MAX_RATE)
    If TEMP_MATRIX(i + 1, 2) < START_RATE Then
        TEMP_MATRIX(i + 1, 4) = PAYMENT_VAL
    Else
        If TEMP_MATRIX(i + 1, 3) = TEMP_MATRIX(i, 3) Then
            TEMP_MATRIX(i + 1, 4) = TEMP_MATRIX(i, 4)
        Else
            TEMP_MATRIX(i + 1, 4) = ANNUITY_FUNC(0, TEMP_MATRIX(i, 8), TEMP_MATRIX(i + 1, 3) / FREQUENCY, j - TEMP_MATRIX(i + 1, 1) + 1)
        End If
    End If
    
    MAX_PAYMENT = MAXIMUM_FUNC(TEMP_MATRIX(i + 1, 4), MAX_PAYMENT)
    If (IsArray(ADDITIONAL_VECTOR) = True) Then
        If i <= UBound(ADDITIONAL_VECTOR, 1) Then
            TEMP_MATRIX(i + 1, 5) = ADDITIONAL_VECTOR(i, 1)
        End If
    End If
    
    TEMP_MATRIX(i + 1, 6) = TEMP_MATRIX(i + 1, 3) / FREQUENCY * TEMP_MATRIX(i, 8)
    TEMP1_SUM = TEMP1_SUM + TEMP_MATRIX(i + 1, 6)
    TEMP_MATRIX(i + 1, 7) = TEMP_MATRIX(i + 1, 4) - TEMP_MATRIX(i + 1, 6) + TEMP_MATRIX(i + 1, 5)
    TEMP2_SUM = TEMP2_SUM + TEMP_MATRIX(i + 1, 7)
    TEMP_MATRIX(i + 1, 8) = TEMP_MATRIX(i, 8) - TEMP_MATRIX(i + 1, 7)
    TEMP_MATRIX(i + 1, 9) = TEMP1_SUM
    TEMP_MATRIX(i + 1, 10) = TEMP2_SUM
Next i
'-----------------------------------------------------------------------------------
Select Case OUTPUT
'-----------------------------------------------------------------------------------
Case 0
'-----------------------------------------------------------------------------------
    ReDim TEMP_VECTOR(1 To 5, 1 To 2)
    
    TEMP_VECTOR(1, 1) = "STARTING MONTHLY PAYMENT"
    TEMP_VECTOR(1, 2) = PAYMENT_VAL
        
    TEMP_VECTOR(2, 1) = "EST. MAX RATE"
    'Based upon the assumptions for "Expected adjustment" and "Periods between
    'adjustments" and the "Interest rate cap", this is the estimated maximum
    'rate that you would expect over the life of the loan.
        
    TEMP_VECTOR(2, 2) = MAX_RATE
    TEMP_VECTOR(3, 1) = "EST. MAX MONTHLY PAYMENT"
    'This is one of the most important estimates! Based upon the
    'rate adjustment assumptions, this is what you'd expect your maximum
    'monthly payment to be over the life of the loan.
    TEMP_VECTOR(3, 2) = MAX_PAYMENT
        
    TEMP_VECTOR(4, 1) = "TOTAL PAYMENTS"
    TEMP_VECTOR(4, 2) = TEMP1_SUM + TEMP2_SUM
        
    TEMP_VECTOR(5, 1) = "TOTAL INTEREST"
    TEMP_VECTOR(5, 2) = TEMP1_SUM
    
    ADJUSTABLE_RATE_MORTGAGE_FUNC = TEMP_VECTOR
'-----------------------------------------------------------------------------------
Case Else
'-----------------------------------------------------------------------------------
    ADJUSTABLE_RATE_MORTGAGE_FUNC = TEMP_MATRIX
'-----------------------------------------------------------------------------------
End Select
'-----------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
ADJUSTABLE_RATE_MORTGAGE_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : Bank's Analysis of Maximum Level of Real Estate Financing
'DESCRIPTION   : DEBT SERVICE FUNCTION
'LIBRARY       : MORTGAGES
'GROUP         : FIXED-INCOME
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'************************************************************************************
'************************************************************************************

Function MORTGAGE_DEBT_SERVICE_FUNC(ByVal SQUARE_FOOTAGE_RNG As Variant, _
ByVal RENTAL_RATE_RNG As Variant, _
ByVal EXPENSES_PERCENTAGE_RNG As Variant, _
ByVal SERVICE_COVERAGE_RNG As Variant, _
ByVal NOMINAL_RATE_RNG As Variant, _
Optional ByVal TENOR_RNG As Variant = 20, _
Optional ByVal FREQUENCY_RNG As Variant = 12)

Dim i As Long
Dim NROWS As Long

Dim SQUARE_FOOTAGE_VECTOR As Variant
Dim RENTAL_RATE_VECTOR As Variant
Dim EXPENSES_PERCENTAGE_VECTOR As Variant
Dim SERVICE_COVERAGE_VECTOR As Variant
Dim NOMINAL_RATE_VECTOR As Variant
Dim TENOR_VECTOR As Variant
Dim FREQUENCY_VECTOR As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

If IsArray(SQUARE_FOOTAGE_RNG) = True Then
    SQUARE_FOOTAGE_VECTOR = SQUARE_FOOTAGE_RNG
    If UBound(SQUARE_FOOTAGE_VECTOR, 1) = 1 Then
        SQUARE_FOOTAGE_VECTOR = MATRIX_TRANSPOSE_FUNC(SQUARE_FOOTAGE_VECTOR)
    End If
Else
    ReDim SQUARE_FOOTAGE_VECTOR(1 To 1, 1 To 1)
    SQUARE_FOOTAGE_VECTOR(1, 1) = SQUARE_FOOTAGE_RNG
End If
NROWS = UBound(SQUARE_FOOTAGE_VECTOR, 1)

If IsArray(RENTAL_RATE_RNG) = True Then
    RENTAL_RATE_VECTOR = RENTAL_RATE_RNG
    If UBound(RENTAL_RATE_VECTOR, 1) = 1 Then
        RENTAL_RATE_VECTOR = MATRIX_TRANSPOSE_FUNC(RENTAL_RATE_VECTOR)
    End If
Else
    ReDim RENTAL_RATE_VECTOR(1 To 1, 1 To 1)
    RENTAL_RATE_VECTOR(1, 1) = RENTAL_RATE_RNG
End If
If UBound(RENTAL_RATE_VECTOR, 1) <> NROWS Then: GoTo ERROR_LABEL

If IsArray(EXPENSES_PERCENTAGE_RNG) = True Then
    EXPENSES_PERCENTAGE_VECTOR = EXPENSES_PERCENTAGE_RNG
    If UBound(EXPENSES_PERCENTAGE_VECTOR, 1) = 1 Then
        EXPENSES_PERCENTAGE_VECTOR = MATRIX_TRANSPOSE_FUNC(EXPENSES_PERCENTAGE_VECTOR)
    End If
Else
    ReDim EXPENSES_PERCENTAGE_VECTOR(1 To 1, 1 To 1)
    EXPENSES_PERCENTAGE_VECTOR(1, 1) = EXPENSES_PERCENTAGE_RNG
End If
If UBound(EXPENSES_PERCENTAGE_VECTOR, 1) <> NROWS Then: GoTo ERROR_LABEL

If IsArray(SERVICE_COVERAGE_RNG) = True Then
    SERVICE_COVERAGE_VECTOR = SERVICE_COVERAGE_RNG
    If UBound(SERVICE_COVERAGE_VECTOR, 1) = 1 Then
        SERVICE_COVERAGE_VECTOR = MATRIX_TRANSPOSE_FUNC(SERVICE_COVERAGE_VECTOR)
    End If
Else
    ReDim SERVICE_COVERAGE_VECTOR(1 To 1, 1 To 1)
    SERVICE_COVERAGE_VECTOR(1, 1) = SERVICE_COVERAGE_RNG
End If
If UBound(SERVICE_COVERAGE_VECTOR, 1) <> NROWS Then: GoTo ERROR_LABEL

If IsArray(NOMINAL_RATE_RNG) = True Then
    NOMINAL_RATE_VECTOR = NOMINAL_RATE_RNG
    If UBound(NOMINAL_RATE_VECTOR, 1) = 1 Then
        NOMINAL_RATE_VECTOR = MATRIX_TRANSPOSE_FUNC(NOMINAL_RATE_VECTOR)
    End If
Else
    ReDim NOMINAL_RATE_VECTOR(1 To 1, 1 To 1)
    NOMINAL_RATE_VECTOR(1, 1) = NOMINAL_RATE_RNG
End If
If UBound(NOMINAL_RATE_VECTOR, 1) <> NROWS Then: GoTo ERROR_LABEL

If IsArray(TENOR_RNG) = True Then
    TENOR_VECTOR = TENOR_RNG
    If UBound(TENOR_VECTOR, 1) = 1 Then
        TENOR_VECTOR = MATRIX_TRANSPOSE_FUNC(TENOR_VECTOR)
    End If
Else
    ReDim TENOR_VECTOR(1 To NROWS, 1 To 1)
    For i = 1 To NROWS
        TENOR_VECTOR(i, 1) = TENOR_RNG
    Next i
End If
If UBound(TENOR_VECTOR, 1) <> NROWS Then: GoTo ERROR_LABEL

If IsArray(FREQUENCY_RNG) = True Then
    FREQUENCY_VECTOR = FREQUENCY_RNG
    If UBound(FREQUENCY_VECTOR, 1) = 1 Then
        FREQUENCY_VECTOR = MATRIX_TRANSPOSE_FUNC(FREQUENCY_VECTOR)
    End If
Else
    ReDim FREQUENCY_VECTOR(1 To NROWS, 1 To 1)
    For i = 1 To NROWS
        FREQUENCY_VECTOR(i, 1) = FREQUENCY_RNG
    Next i
End If
If UBound(FREQUENCY_VECTOR, 1) <> NROWS Then: GoTo ERROR_LABEL

ReDim TEMP_MATRIX(0 To NROWS, 1 To 10)

TEMP_MATRIX(0, 1) = "SQUARE FOOTAGE"
TEMP_MATRIX(0, 2) = "RENTAL RATE PER SQUARE FOOT"
TEMP_MATRIX(0, 3) = "TOTAL GROSS RENTAL INCOME"
TEMP_MATRIX(0, 4) = "LESS: VACANCY, MAINTENANCE FEE AND CREDIT LOSS"
TEMP_MATRIX(0, 5) = "DEBT SERVICE CAPABILITY"
TEMP_MATRIX(0, 6) = "NET OPERATING INCOME (BEFORE DEBT SERVICE) / MINIMUM DEBT SERVICE COVERAGE"
TEMP_MATRIX(0, 7) = "ON A COUNT BASIS"
TEMP_MATRIX(0, 8) = "NOMINAL ANNUAL RATE"
TEMP_MATRIX(0, 9) = "EFFECTIVE PERIOD INTEREST RATE"
TEMP_MATRIX(0, 10) = "MAXIMUM FINANCING BASED ON DEBT SERVICE CAPABILITY"

For i = 1 To NROWS
    TEMP_MATRIX(i, 1) = SQUARE_FOOTAGE_VECTOR(i, 1)
    TEMP_MATRIX(i, 2) = RENTAL_RATE_VECTOR(i, 1)
    TEMP_MATRIX(i, 3) = TEMP_MATRIX(i, 1) * TEMP_MATRIX(i, 2)
    TEMP_MATRIX(i, 4) = TEMP_MATRIX(i, 3) * EXPENSES_PERCENTAGE_VECTOR(i, 1)
    TEMP_MATRIX(i, 5) = TEMP_MATRIX(i, 3) - TEMP_MATRIX(i, 4)
    TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 5) / SERVICE_COVERAGE_VECTOR(i, 1)
    TEMP_MATRIX(i, 7) = TEMP_MATRIX(i, 6) / FREQUENCY_VECTOR(i, 1)
    TEMP_MATRIX(i, 8) = NOMINAL_RATE_VECTOR(i, 1)
    TEMP_MATRIX(i, 9) = (1 + TEMP_MATRIX(i, 8) / 2) ^ (1 / (FREQUENCY / 2)) - 1
    TEMP_MATRIX(i, 10) = TEMP_MATRIX(i, 7) * (1 - (1 / (1 + TEMP_MATRIX(i, 9))) ^ (TENOR_VECTOR(i, 1) * FREQUENCY_VECTOR(i, 1))) / TEMP_MATRIX(i, 9)
Next i

MORTGAGE_DEBT_SERVICE_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
MORTGAGE_DEBT_SERVICE_FUNC = Err.number
End Function
