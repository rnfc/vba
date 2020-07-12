Attribute VB_Name = "FINAN_FI_MBS_PAYM_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : MBS_CMO_FUNC

'DESCRIPTION   : A CMO structure, also known as Real Estate Mortgage
'Investment Conduits (REMICS), is a mechanism for reallocating cash flows
'from one or more mortgage pass-through or a pool of mortgages into multiple
'classes with different priority claims.

'LIBRARY       : MBS
'GROUP         : PAYMENT
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'************************************************************************************
'************************************************************************************

Function MBS_CMO_FUNC(ByVal FIRST_PAYMENT As Date, _
ByVal TENOR As Double, _
ByVal POOL_SIZE As Double, _
ByVal WACC_VAL As Double, _
ByVal PASSTHROUGH_RATE As Double, _
ByVal CPR_VAL As Double, _
ByVal ADJUSTED_PERIODS As Double, _
ByRef TRANCHE_PAR_RNG As Variant, _
ByRef TRANCHE_COUP_RNG As Variant, _
Optional ByVal FREQUENCY As Long = 12, _
Optional ByVal FACTOR As Double = 1.65, _
Optional ByVal SEASONING As Double = 3, _
Optional ByVal OUTPUT As Integer = 1)

'Payment Rules:

'* Coupon Interest:  Disburse periodic coupon interest to each tranche
'on the basis of principal outstanding at the beginning of that period.

'* Principal:  First to Tranche A until completely paid off.
'                Next to Tranch B until completely paid off.
'                Once B is paid off, disburse to C until it is
'                   fully paid off.
'                Once C is paid off, disburse to D until it is
'                   fully paid off.......................

Dim i As Long '
Dim j As Long
Dim k As Long
Dim l As Long

Dim NTRANCHES As Long
Dim TEMP_STR As String

Dim PASS_MATRIX As Variant
Dim TEMP_MATRIX As Variant

Dim PAR_VECTOR As Variant
Dim COUPON_VECTOR As Variant
Dim TEMP_VECTOR As Variant

Dim tolerance As Double

On Error GoTo ERROR_LABEL

PAR_VECTOR = TRANCHE_PAR_RNG
If UBound(PAR_VECTOR, 1) = 1 Then
    PAR_VECTOR = MATRIX_TRANSPOSE_FUNC(PAR_VECTOR)
End If

COUPON_VECTOR = TRANCHE_COUP_RNG
If UBound(COUPON_VECTOR, 1) = 1 Then
    COUPON_VECTOR = MATRIX_TRANSPOSE_FUNC(COUPON_VECTOR)
End If

If UBound(PAR_VECTOR, 1) <> UBound(COUPON_VECTOR, 1) Then: GoTo ERROR_LABEL

NTRANCHES = UBound(PAR_VECTOR, 1)
PASS_MATRIX = MBS_PASS_FUNC(FIRST_PAYMENT, TENOR, POOL_SIZE, WACC_VAL, _
           PASSTHROUGH_RATE, CPR_VAL, ADJUSTED_PERIODS, FREQUENCY, _
           FACTOR, SEASONING)

k = 6
ReDim TEMP_MATRIX(0 To UBound(PASS_MATRIX, 1), 1 To ((NTRANCHES * k) + 2))

'------------------------------First Pass: Setting Periods
TEMP_MATRIX(0, 1) = "PERIODS"
TEMP_MATRIX(0, 2) = "MATURITY"

For i = 1 To UBound(PASS_MATRIX, 1)
    TEMP_MATRIX(i, 1) = PASS_MATRIX(i, 1)
    TEMP_MATRIX(i, 2) = PASS_MATRIX(i, 2)
Next i

'------------------------------Second Pass: Setting Table

j = 3
tolerance = 0.0001
l = 0

Do Until l = NTRANCHES
    
    l = l + 1
    TEMP_STR = "TRANCHE #" & l
        
        TEMP_MATRIX(0, j) = TEMP_STR & " : BALANCE"
        TEMP_MATRIX(0, j + 1) = TEMP_STR & " : PRINCIPAL"
        TEMP_MATRIX(0, j + 2) = TEMP_STR & " : INTEREST"
        TEMP_MATRIX(0, j + 3) = TEMP_STR & " : CUMULATIVE" 'PRINCIPAL
        TEMP_MATRIX(0, j + 4) = TEMP_STR & " : WEIGHTED" 'Weighted Principal Pmts
        TEMP_MATRIX(0, j + 5) = TEMP_STR & " : PV CFs" 'PV Cash Flows
        
        TEMP_MATRIX(1, j) = PAR_VECTOR(l, 1)
        
        If l = 1 Then
            TEMP_MATRIX(1, j + 1) = PASS_MATRIX(1, 9) + PASS_MATRIX(1, 10)
        Else
            If (TEMP_MATRIX(1, (j) - k) - TEMP_MATRIX(1, (j + 1) - k)) < tolerance Then
                TEMP_MATRIX(1, j + 1) = MINIMUM_FUNC(PASS_MATRIX(1, 9) + _
                PASS_MATRIX(1, 10) - TEMP_MATRIX(1, (j + 1) - k), _
                PAR_VECTOR(l, 1) - 0)
            Else
                TEMP_MATRIX(1, j + 1) = 0
            End If
        End If
        
        TEMP_MATRIX(1, j + 2) = TEMP_MATRIX(1, j) * _
        COUPON_VECTOR(l, 1) / FREQUENCY
        
        TEMP_MATRIX(1, j + 3) = TEMP_MATRIX(1, j + 1)
                
        TEMP_MATRIX(1, j + 4) = TEMP_MATRIX(1, j + 1) * TEMP_MATRIX(1, 1)
        TEMP_MATRIX(1, j + 5) = (TEMP_MATRIX(1, j + 1) + TEMP_MATRIX(1, j + 2)) _
        / (1 + COUPON_VECTOR(l, 1) / FREQUENCY) ^ TEMP_MATRIX(1, 1)
                
        For i = 2 To UBound(PASS_MATRIX, 1)
            TEMP_MATRIX(i, j) = TEMP_MATRIX(i - 1, j) - TEMP_MATRIX(i - 1, j + 1)
            
            If l = 1 Then
                TEMP_MATRIX(i, j + 1) = MINIMUM_FUNC(PASS_MATRIX(i, 9) + _
                PASS_MATRIX(i, 10), PAR_VECTOR(l, 1) - TEMP_MATRIX(i - 1, j + 3))
            Else
            
                If (TEMP_MATRIX(i, (j) - k) - TEMP_MATRIX(i, (j + 1) - k)) < tolerance Then
                    TEMP_MATRIX(i, j + 1) = MINIMUM_FUNC(PASS_MATRIX(i, 9) + _
                    PASS_MATRIX(i, 10) - TEMP_MATRIX(i, (j + 1) - k), _
                    PAR_VECTOR(l, 1) - TEMP_MATRIX(i - 1, j + 3))
                Else
                    TEMP_MATRIX(i, j + 1) = 0
                End If
            End If
            
            TEMP_MATRIX(i, j + 2) = TEMP_MATRIX(i, j) * _
            COUPON_VECTOR(l, 1) / FREQUENCY
            
            TEMP_MATRIX(i, j + 3) = TEMP_MATRIX(i, j + 1) + TEMP_MATRIX(i - 1, j + 3)
            
            TEMP_MATRIX(i, j + 4) = TEMP_MATRIX(i, j + 1) * TEMP_MATRIX(i, 1)
        
            TEMP_MATRIX(i, j + 5) = (TEMP_MATRIX(i, j + 1) + TEMP_MATRIX(i, j + 2)) _
            / (1 + COUPON_VECTOR(l, 1) / FREQUENCY) ^ TEMP_MATRIX(i, 1)
        Next i
    
    j = j + k
Loop


Select Case OUTPUT
Case 0
    MBS_CMO_FUNC = TEMP_MATRIX
Case Else
    ReDim TEMP_VECTOR(0 To NTRANCHES, 1 To 4)
    TEMP_VECTOR(0, 1) = "-"
    TEMP_VECTOR(0, 2) = "PAYMENT"
    TEMP_VECTOR(0, 3) = "AVG_LIFE"
    TEMP_VECTOR(0, 4) = "NPV"
    j = 3
    For i = 1 To NTRANCHES
        
        TEMP_VECTOR(i, 1) = "TRANCHE #" & i
        TEMP_VECTOR(i, 2) = ANNUITY_FUNC(0, PAR_VECTOR(i, 1), _
            WACC_VAL / FREQUENCY, UBound(PASS_MATRIX, 1))
        
        TEMP_VECTOR(i, 3) = _
        MATRIX_ELEMENTS_CUMULATIVE_SUM_FUNC(MATRIX_REMOVE_ROWS_FUNC(MATRIX_GET_COLUMN_FUNC( _
        TEMP_MATRIX, j + 4, 1), 1, 1)) / (FREQUENCY * PAR_VECTOR(i, 1))
        
        TEMP_VECTOR(i, 4) = _
        MATRIX_ELEMENTS_CUMULATIVE_SUM_FUNC(MATRIX_REMOVE_ROWS_FUNC(MATRIX_GET_COLUMN_FUNC( _
        TEMP_MATRIX, j + 5, 1), 1, 1)) - PAR_VECTOR(i, 1)
        j = j + k
    Next i
    MBS_CMO_FUNC = TEMP_VECTOR
End Select

Exit Function
ERROR_LABEL:
MBS_CMO_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MBS_PSA_FUNC
'DESCRIPTION   : Table of Cash Flows from the mortgage pool -- using PSA
'Benchmark (with indicated Factor)
'LIBRARY       : MBS
'GROUP         : PAYMENT
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'************************************************************************************
'************************************************************************************

Function MBS_PSA_FUNC(ByVal FIRST_PAYMENT As Date, _
ByVal TENOR As Double, _
ByVal LOAN_AMOUNT As Double, _
ByVal RATE_VAL As Double, _
ByVal CPR_VAL As Double, _
ByVal ADJUSTED_PERIODS As Double, _
Optional ByVal FREQUENCY As Long = 12, _
Optional ByVal FACTOR As Double = 1)

'-------------------------ADJUSTED_PERIODS: In Years

'------------------------Wall Street 's obsessions-------------------------
'Historically in the US FRMs have imbedded in the interest RATE_VAL
'inflationary expectations as well as any inflation risk premia.

'This causes a "tilt" or shift in the timing of the real cash flows
'to the earlier days of the mortgage.

'This, in turn can reduce the "affordability" of housing since banks
'rely on ratios of monthly payments to income for qualification.


'Prepayment is an option given to the borrower to put the
'loan at par to the lender. Unlike exchange traded options,
'this is not necessarily exercised if and only if market
'conditions make it optimal to do so.
'    Reasons for prepayment:
'        1) Refinancing
'            a) Lower interest rates.
'            b) Extracting equity from asset.
'        2) Moving  (and loan not assumed).

'--------------------------------------------------------------------------

'Public Securities Association Convention (Benchmark)
'Standard prepayment assumption model of The Bond Market Association.

'    CPR_VAL = 6% (t/30) when 1 < t < 30.
'    CPR_VAL = 6%  when t > 30.
'    CPR_VAL: Conditional Prepayment Rate

'SMM: Single Monthly Mortality RATE_VAL.

Dim i As Long
Dim j As Long

Dim PAYMENT_VAL As Double
Dim ADJ_PERIOD As Double

Dim TEMP_MATRIX As Variant


On Error GoTo ERROR_LABEL

j = Int(FREQUENCY * TENOR)
PAYMENT_VAL = ANNUITY_FUNC(0, LOAN_AMOUNT, RATE_VAL / FREQUENCY, j)
ADJ_PERIOD = ADJUSTED_PERIODS * FREQUENCY

ReDim TEMP_MATRIX(0 To j, 1 To 11)

TEMP_MATRIX(0, 1) = "PERIODS"
TEMP_MATRIX(0, 2) = "MATURITY"
TEMP_MATRIX(0, 3) = "CPR_VAL"
TEMP_MATRIX(0, 4) = "SSM"
TEMP_MATRIX(0, 5) = "PRODUCT: (1-SMM)"
TEMP_MATRIX(0, 6) = "PAYMENT"
TEMP_MATRIX(0, 7) = "INTEREST"
TEMP_MATRIX(0, 8) = "PRINCIPAL"
TEMP_MATRIX(0, 9) = "PRE-PAYMENT"
TEMP_MATRIX(0, 10) = "REMAINING"
TEMP_MATRIX(0, 11) = "CASH FLOW"

TEMP_MATRIX(1, 1) = 1
TEMP_MATRIX(1, 2) = EDATE_FUNC(FIRST_PAYMENT, TEMP_MATRIX(1, 1) * 12 / FREQUENCY)

If TEMP_MATRIX(1, 1) > ADJ_PERIOD Then
    TEMP_MATRIX(1, 3) = CPR_VAL * FACTOR
Else
    TEMP_MATRIX(1, 3) = (TEMP_MATRIX(1, 1) / ADJ_PERIOD) * CPR_VAL * FACTOR
End If

TEMP_MATRIX(1, 4) = 1 - (1 - TEMP_MATRIX(1, 3)) ^ (1 / FREQUENCY)
TEMP_MATRIX(1, 5) = 1 - TEMP_MATRIX(1, 4)
TEMP_MATRIX(1, 6) = PAYMENT_VAL
TEMP_MATRIX(1, 7) = LOAN_AMOUNT / FREQUENCY * RATE_VAL
TEMP_MATRIX(1, 8) = PAYMENT_VAL - TEMP_MATRIX(1, 7)
TEMP_MATRIX(1, 9) = TEMP_MATRIX(1, 4) * (LOAN_AMOUNT - TEMP_MATRIX(1, 8))

TEMP_MATRIX(1, 10) = LOAN_AMOUNT - TEMP_MATRIX(1, 8) - TEMP_MATRIX(1, 9)
TEMP_MATRIX(1, 11) = TEMP_MATRIX(1, 6) + TEMP_MATRIX(1, 9)

For i = 2 To j
    
    TEMP_MATRIX(i, 1) = i
    TEMP_MATRIX(i, 2) = EDATE_FUNC(FIRST_PAYMENT, TEMP_MATRIX(i, 1) * 12 / FREQUENCY)
    
    If TEMP_MATRIX(i, 1) > ADJ_PERIOD Then
        TEMP_MATRIX(i, 3) = CPR_VAL * FACTOR
    Else
        TEMP_MATRIX(i, 3) = (TEMP_MATRIX(i, 1) / ADJ_PERIOD) * CPR_VAL * FACTOR
    End If
    
    TEMP_MATRIX(i, 4) = 1 - (1 - TEMP_MATRIX(i, 3)) ^ (1 / FREQUENCY)
    TEMP_MATRIX(i, 5) = (1 - TEMP_MATRIX(i, 4)) * TEMP_MATRIX(i - 1, 5)
    TEMP_MATRIX(i, 6) = PAYMENT_VAL * TEMP_MATRIX(i - 1, 5)
    TEMP_MATRIX(i, 7) = TEMP_MATRIX(i - 1, 10) / FREQUENCY * RATE_VAL
    TEMP_MATRIX(i, 8) = TEMP_MATRIX(i, 6) - TEMP_MATRIX(i, 7)
    
    TEMP_MATRIX(i, 9) = TEMP_MATRIX(i, 4) * (TEMP_MATRIX(i - 1, 10) - TEMP_MATRIX(i, 8))
    
    TEMP_MATRIX(i, 10) = TEMP_MATRIX(i - 1, 10) - TEMP_MATRIX(i, 8) - TEMP_MATRIX(i, 9)
    TEMP_MATRIX(i, 11) = TEMP_MATRIX(i, 6) + TEMP_MATRIX(i, 9)
Next i

MBS_PSA_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
MBS_PSA_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MBS_PASS_FUNC
'DESCRIPTION   : Pass through Mortgage Backed Security

'In a PASS THROUGH MBS the originator keeps a servicing fee, but
'otherwise, as the name implies, all cash flows flow through to
'the bond holder.

'GNMA provides a guarantee against default on its pass-throughs
'(which is backed by the full faith and credit of the US government).

'LIBRARY       : MBS
'GROUP         : PAYMENT
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'************************************************************************************
'************************************************************************************


Function MBS_PASS_FUNC(ByVal FIRST_PAYMENT As Date, _
ByVal TENOR As Double, _
ByVal POOL_SIZE As Double, _
ByVal WACC_VAL As Double, _
ByVal PASSTHROUGH_RATE As Double, _
ByVal CPR_VAL As Double, _
ByVal ADJUSTED_PERIODS As Double, _
Optional ByVal FREQUENCY As Long = 12, _
Optional ByVal FACTOR As Double = 1.65, _
Optional ByVal SEASONING As Double = 3)

'TENOR: DO NOT include the implied seasoning periods
'WACC_VAL: Weighted Average Coupon (WAC)
'(TENOR * FREQUENCY) + SEASONING = WAM: Weighted Average Maturity in periods
'ADJUSTED_PERIODS: In Years
'SEASONING: In Periods

'POOL_SIZE:  400,000,000.00
'TENOR: 29.5
'ADJUSTED_PERIODS: 2.25
'START_RATE:  8.13%
'PASSTHROUGH_RATE:    7.50%
'CPR_VAL: 6.00%
'FIRST_PAYMENT: 15 / 11 / 2006

Dim i As Long
Dim j As Double 'Wam :Weighted Average Maturity in periods

Dim ADJ_PERIOD As Double
Dim PAYMENT_VAL As Double
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

j = Int(FREQUENCY * TENOR) + SEASONING
PAYMENT_VAL = ANNUITY_FUNC(0, POOL_SIZE, WACC_VAL / FREQUENCY, j)
ADJ_PERIOD = ADJUSTED_PERIODS * FREQUENCY

ReDim TEMP_MATRIX(0 To j, 1 To 13)

TEMP_MATRIX(0, 1) = "PERIODS"
TEMP_MATRIX(0, 2) = "MATURITY"
TEMP_MATRIX(0, 3) = "CPR_VAL"
TEMP_MATRIX(0, 4) = "SSM"
TEMP_MATRIX(0, 5) = "PRODUCT: (1-SMM)"
TEMP_MATRIX(0, 6) = "PAYMENT"
TEMP_MATRIX(0, 7) = "INTEREST"
TEMP_MATRIX(0, 8) = "NET INTEREST"
TEMP_MATRIX(0, 9) = "PRINCIPAL"
TEMP_MATRIX(0, 10) = "PRE-PAYMENT"
TEMP_MATRIX(0, 11) = "REMAINING"
TEMP_MATRIX(0, 12) = "CASH FLOW" 'PASSTHROUGH CASH FLOW
TEMP_MATRIX(0, 13) = "WEIGHTED PRINCIPAL" 'TIME-WEIGHTED PRINCIPAL PAYMENT

TEMP_MATRIX(1, 1) = 1
TEMP_MATRIX(1, 2) = EDATE_FUNC(FIRST_PAYMENT, TEMP_MATRIX(1, 1) * 12 / FREQUENCY)

If TEMP_MATRIX(1, 1) > ADJ_PERIOD Then
    TEMP_MATRIX(1, 3) = MINIMUM_FUNC(CPR_VAL * FACTOR, 1)
Else
    TEMP_MATRIX(1, 3) = MINIMUM_FUNC(((TEMP_MATRIX(1, 1) + _
    SEASONING) / (ADJ_PERIOD + SEASONING)) * CPR_VAL * FACTOR, 1)
End If

TEMP_MATRIX(1, 4) = 1 - (1 - TEMP_MATRIX(1, 3)) ^ (1 / FREQUENCY)
TEMP_MATRIX(1, 5) = 1 - TEMP_MATRIX(1, 4)
TEMP_MATRIX(1, 6) = PAYMENT_VAL
TEMP_MATRIX(1, 7) = POOL_SIZE / FREQUENCY * WACC_VAL
TEMP_MATRIX(1, 8) = POOL_SIZE / FREQUENCY * PASSTHROUGH_RATE

TEMP_MATRIX(1, 9) = TEMP_MATRIX(1, 6) - TEMP_MATRIX(1, 7)
TEMP_MATRIX(1, 10) = TEMP_MATRIX(1, 4) * (POOL_SIZE - TEMP_MATRIX(1, 9))

TEMP_MATRIX(1, 11) = POOL_SIZE - TEMP_MATRIX(1, 9) - TEMP_MATRIX(1, 10)

TEMP_MATRIX(1, 12) = TEMP_MATRIX(1, 8) + TEMP_MATRIX(1, 9) + TEMP_MATRIX(1, 10)
TEMP_MATRIX(1, 13) = TEMP_MATRIX(1, 1) * (TEMP_MATRIX(1, 9) + TEMP_MATRIX(1, 10))

For i = 2 To j
    
    TEMP_MATRIX(i, 1) = i
    TEMP_MATRIX(i, 2) = EDATE_FUNC(FIRST_PAYMENT, TEMP_MATRIX(i, 1) * 12 / FREQUENCY)
    
    If TEMP_MATRIX(i, 1) > ADJ_PERIOD Then
        TEMP_MATRIX(i, 3) = MINIMUM_FUNC(CPR_VAL * FACTOR, 1)
    Else
        TEMP_MATRIX(i, 3) = MINIMUM_FUNC(((TEMP_MATRIX(i, 1) + _
        SEASONING) / (ADJ_PERIOD + SEASONING)) * CPR_VAL * FACTOR, 1)
    End If
    
    TEMP_MATRIX(i, 4) = 1 - (1 - TEMP_MATRIX(i, 3)) ^ (1 / FREQUENCY)
    TEMP_MATRIX(i, 5) = (1 - TEMP_MATRIX(i, 4)) * TEMP_MATRIX(i - 1, 5)
    TEMP_MATRIX(i, 6) = PAYMENT_VAL * TEMP_MATRIX(i - 1, 5)
    TEMP_MATRIX(i, 7) = TEMP_MATRIX(i - 1, 11) / FREQUENCY * WACC_VAL
    TEMP_MATRIX(i, 8) = TEMP_MATRIX(i - 1, 11) / FREQUENCY * PASSTHROUGH_RATE
    TEMP_MATRIX(i, 9) = TEMP_MATRIX(i, 6) - TEMP_MATRIX(i, 7)
    TEMP_MATRIX(i, 10) = TEMP_MATRIX(i, 4) * (TEMP_MATRIX(i - 1, 11) - TEMP_MATRIX(i, 9))
    
    TEMP_MATRIX(i, 11) = TEMP_MATRIX(i - 1, 11) - TEMP_MATRIX(i, 9) - TEMP_MATRIX(i, 10)
    TEMP_MATRIX(i, 12) = TEMP_MATRIX(i, 8) + TEMP_MATRIX(i, 9) + TEMP_MATRIX(i, 10)
    TEMP_MATRIX(i, 13) = TEMP_MATRIX(i, 1) * (TEMP_MATRIX(i, 9) + TEMP_MATRIX(i, 10))
Next i

MBS_PASS_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
MBS_PASS_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MBS_PAC_FUNC
'DESCRIPTION   : Planned Amortization Class NTRANCHES

'History of PAC: In 1986, following large drops in mortgage rates, and
'increased pre-payments, issuers developed and issued prepayment protected
'bonds called planned amortization classes.  (PACs)

'PACs have a principal payment schedule--like a sinking fund that can be
'maintained over a range of prepayment rates.

'The relative certainty of these PAC NTRANCHES is obtained by setting up
'a companion tranche that absorbs the uncertainty. Greater Certainty
'achieved by shifting the uncertainty to a support or companion bond.

'LIBRARY       : MBS
'GROUP         : PAYMENT
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'************************************************************************************
'************************************************************************************


Function MBS_PAC_FUNC(ByVal FIRST_PAYMENT As Date, _
ByVal TENOR As Double, _
ByVal POOL_SIZE As Double, _
ByVal WACC_VAL As Double, _
ByVal PASSTHROUGH_RATE As Double, _
ByVal CPR_VAL As Double, _
ByVal ADJUSTED_PERIODS As Double, _
ByRef TRANCHE_PAR_RNG As Variant, _
ByRef TRANCHE_COUP_RNG As Variant, _
ByRef TRANCHE_FACTOR_RNG As Variant, _
Optional ByVal FREQUENCY As Long = 12, _
Optional ByVal SEASONING As Double = 3)

'Payment Rules:
'Coupon Interest:  Disburse periodic coupon interest to each tranch on the
'basis of principal outstanding at the beginning of that period.

'Principal: Disburse to PAC based on its schedule.  Any excess pay to support.
'When S is fuly paid off, pay to P, regardless of schedule.

'Note that in the event that the realized prepayment rates lie
'between 90% and 300% PSA rates, the cash flows to the
'PAC are the same.

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim NTRANCHES As Long

Dim TEMP_STR As String

Dim TEMP_MATRIX As Variant
Dim PASS_MATRIX As Variant

Dim PAR_VECTOR As Variant
Dim COUPON_VECTOR As Variant
Dim FACTOR_VECTOR As Variant

Dim TEMP_MIN As Variant

On Error GoTo ERROR_LABEL

PAR_VECTOR = TRANCHE_PAR_RNG
If UBound(PAR_VECTOR, 1) = 1 Then
    PAR_VECTOR = MATRIX_TRANSPOSE_FUNC(PAR_VECTOR)
End If

COUPON_VECTOR = TRANCHE_COUP_RNG
If UBound(COUPON_VECTOR, 1) = 1 Then
    COUPON_VECTOR = MATRIX_TRANSPOSE_FUNC(COUPON_VECTOR)
End If
If UBound(PAR_VECTOR, 1) <> UBound(COUPON_VECTOR, 1) Then: GoTo ERROR_LABEL

FACTOR_VECTOR = TRANCHE_FACTOR_RNG
If UBound(FACTOR_VECTOR, 1) = 1 Then
    FACTOR_VECTOR = MATRIX_TRANSPOSE_FUNC(FACTOR_VECTOR)
End If
If UBound(PAR_VECTOR, 1) <> UBound(FACTOR_VECTOR, 1) Then: GoTo ERROR_LABEL

NTRANCHES = UBound(PAR_VECTOR, 1)
NROWS = Int(FREQUENCY * TENOR) + SEASONING
ReDim TEMP_MATRIX(0 To NROWS, 1 To (NTRANCHES + 3))

'------------------------------First Pass: Setting Periods
TEMP_MATRIX(0, 1) = "PERIODS"
TEMP_MATRIX(0, 2) = "MATURITY"

For i = 1 To NROWS
    TEMP_MATRIX(i, 1) = i
    TEMP_MATRIX(i, 2) = EDATE_FUNC(FIRST_PAYMENT, _
    TEMP_MATRIX(i, 1) * 12 / FREQUENCY)
Next i

'------------------------------Second Pass: Setting Table

For j = 1 To NTRANCHES
    TEMP_STR = "PRINCIPAL: " & "AT " & Format(FACTOR_VECTOR(j, 1), "0.0%") & " PSA"
    TEMP_MATRIX(0, j + 2) = TEMP_STR
    PASS_MATRIX = MBS_PASS_FUNC(FIRST_PAYMENT, TENOR, POOL_SIZE, WACC_VAL, PASSTHROUGH_RATE, CPR_VAL, ADJUSTED_PERIODS, FREQUENCY, FACTOR_VECTOR(j, 1), SEASONING)
    For i = 1 To NROWS
        TEMP_MATRIX(i, j + 2) = PASS_MATRIX(i, 9) + PASS_MATRIX(i, 10)
    Next i
Next j

TEMP_MATRIX(0, NTRANCHES + 3) = "MINIMUM: PAC SCHEDULE"
For i = 1 To NROWS
    TEMP_MIN = TEMP_MATRIX(i, 2 + 1)
    For j = 2 To NTRANCHES
        If TEMP_MATRIX(i, j + 2) < TEMP_MIN Then: TEMP_MIN = TEMP_MATRIX(i, j + 2)
    Next j
    TEMP_MATRIX(i, NTRANCHES + 3) = TEMP_MIN
Next i
                
MBS_PAC_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
MBS_PAC_FUNC = Err.number
End Function
