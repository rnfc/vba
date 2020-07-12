Attribute VB_Name = "FINAN_FI_BOND_SWAP_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : SWAP_COMPARATIVE_ADVANTAGE_FUNC
'DESCRIPTION   : YIELD THE COMPARATIVE ADVANTAGE IN A SWAP TRANSACTION BETWEEN
'TWO PARTIES. FLOATING RATES MUST BE QUOTED IN BASIS POINTS
'FIXED RATES MUST BE QUOTED IN PERCENTAGES

'LIBRARY       : SWAP
'GROUP         : ARBITRAGE
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'**********************************************************************************
'**********************************************************************************

Function SWAP_COMPARATIVE_ADVANTAGE_FUNC(ByVal CLIENT1_STR As String, _
ByVal CLIENT2_STR As String, _
ByVal FIXED1_RATE As Variant, _
ByVal FIXED2_RATE As Variant, _
ByVal FLOATING1_RATE As Variant, _
ByVal FLOATING2_RATE As Variant)

Dim FIXED_STR As String
Dim FLOATING_STR As String

On Error GoTo ERROR_LABEL

If Abs(Val(FIXED2_RATE) - Val(FIXED1_RATE)) = Abs(Val(FLOATING2_RATE) / 10000 - Val(FLOATING1_RATE) / 10000) Then
    FIXED_STR = "-"
ElseIf Abs(Val(FIXED2_RATE) - Val(FIXED1_RATE)) > Abs(Val(FLOATING2_RATE) / 10000 - Val(FLOATING1_RATE) / 10000) Then
    FIXED_STR = CLIENT1_STR
Else
    FIXED_STR = CLIENT2_STR
End If

If Abs(Val(FIXED2_RATE) - Val(FIXED1_RATE)) = Abs(Val(FLOATING2_RATE) / 10000 - Val(FLOATING1_RATE) / 10000) Then
    FLOATING_STR = "-"
ElseIf Abs(Val(FLOATING2_RATE) / 10000 - Val(FLOATING1_RATE) / 10000) > Abs(Val(FIXED2_RATE) - Val(FIXED1_RATE)) Then
    FLOATING_STR = CLIENT1_STR
Else
    FLOATING_STR = CLIENT2_STR
End If

'Comparative Advantage
SWAP_COMPARATIVE_ADVANTAGE_FUNC = Array(FIXED_STR, FLOATING_STR)

Exit Function
ERROR_LABEL:
SWAP_COMPARATIVE_ADVANTAGE_FUNC = Err.number
End Function


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : SWAP_ARBITRAGE_FUNC
'DESCRIPTION   : Plain Vanilla Swap Arbitrage Model
'LIBRARY       : SWAP
'GROUP         : ARBITRAGE
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'**********************************************************************************
'**********************************************************************************

Function SWAP_ARBITRAGE_FUNC(ByVal NOTIONAL As Double, _
ByVal FIXED_RATE_CLIENT As Double, _
ByVal RECEIVES_BPS As Double, _
ByVal START_DATE As Date, _
ByVal END_DATE As Date, _
ByVal FREQUENCY As Integer, _
ByRef DATES_RNG As Variant, _
ByRef DISCOUNT_RNG As Variant, _
Optional ByVal OUTPUT As Integer = 0, _
Optional ByVal COUNT_BASIS As Integer = 3, _
Optional ByVal INTER_OPT As Integer = 0)

Dim i As Long
Dim j As Long

Dim TEMP1_SUM As Double
Dim TEMP2_SUM As Double

Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

j = COUPNUM_FUNC(START_DATE, END_DATE, FREQUENCY)

ReDim TEMP_MATRIX(0 To j, 1 To 12)
TEMP_MATRIX(0, 1) = "START DATE"
TEMP_MATRIX(0, 2) = "END DATE"
TEMP_MATRIX(0, 3) = "NOTIONAL"
TEMP_MATRIX(0, 4) = "DC"
TEMP_MATRIX(0, 5) = "PAYMENT RATE"
TEMP_MATRIX(0, 6) = "PAYMENT"
TEMP_MATRIX(0, 7) = "DF AT START"
TEMP_MATRIX(0, 8) = "DF AT END"
TEMP_MATRIX(0, 9) = "PV OF PAYMENT"
TEMP_MATRIX(0, 10) = "PAYMENT RATE"
TEMP_MATRIX(0, 11) = "PAYMENT"
TEMP_MATRIX(0, 12) = "PAYMENT PV OF PAYMENT"

TEMP_MATRIX(1, 1) = START_DATE
TEMP_MATRIX(1, 2) = EDATE_FUNC(TEMP_MATRIX(1, 1), 12 / FREQUENCY)
    
If TEMP_MATRIX(1, 2) <= END_DATE Then: TEMP_MATRIX(1, 3) = NOTIONAL
    
TEMP_MATRIX(1, 4) = YEARFRAC_FUNC(TEMP_MATRIX(1, 1), TEMP_MATRIX(1, 2), COUNT_BASIS)
TEMP_MATRIX(1, 5) = FIXED_RATE_CLIENT
TEMP_MATRIX(1, 6) = -TEMP_MATRIX(1, 3) * TEMP_MATRIX(1, 4) * TEMP_MATRIX(1, 5)
    
TEMP_MATRIX(1, 7) = YIELD_INTERPOLATION_FUNC(TEMP_MATRIX(1, 1), DATES_RNG, DISCOUNT_RNG, INTER_OPT, 0)
TEMP_MATRIX(1, 8) = YIELD_INTERPOLATION_FUNC(TEMP_MATRIX(1, 2), DATES_RNG, DISCOUNT_RNG, INTER_OPT, 0)
TEMP_MATRIX(1, 9) = TEMP_MATRIX(1, 8) * TEMP_MATRIX(1, 6)
TEMP_MATRIX(1, 10) = (TEMP_MATRIX(1, 7) / TEMP_MATRIX(1, 8) - 1) / TEMP_MATRIX(1, 4) + RECEIVES_BPS / 10000
TEMP_MATRIX(1, 11) = TEMP_MATRIX(1, 3) * TEMP_MATRIX(1, 10) * TEMP_MATRIX(1, 4)
TEMP_MATRIX(1, 12) = TEMP_MATRIX(1, 11) * TEMP_MATRIX(1, 8)
    
TEMP1_SUM = TEMP_MATRIX(1, 9)
TEMP2_SUM = TEMP_MATRIX(1, 12)

For i = 2 To j
    TEMP_MATRIX(i, 1) = TEMP_MATRIX(i - 1, 2)
    TEMP_MATRIX(i, 2) = EDATE_FUNC(TEMP_MATRIX(i, 1), 12 / FREQUENCY)
    If TEMP_MATRIX(i, 2) <= END_DATE Then: TEMP_MATRIX(i, 3) = NOTIONAL
    TEMP_MATRIX(i, 4) = YEARFRAC_FUNC(TEMP_MATRIX(i, 1), TEMP_MATRIX(i, 2), COUNT_BASIS)
    TEMP_MATRIX(i, 5) = FIXED_RATE_CLIENT
    TEMP_MATRIX(i, 6) = -TEMP_MATRIX(i, 3) * TEMP_MATRIX(i, 4) * TEMP_MATRIX(i, 5)
    TEMP_MATRIX(i, 7) = YIELD_INTERPOLATION_FUNC(TEMP_MATRIX(i, 1), DATES_RNG, DISCOUNT_RNG, INTER_OPT, 0)
    TEMP_MATRIX(i, 8) = YIELD_INTERPOLATION_FUNC(TEMP_MATRIX(i, 2), DATES_RNG, DISCOUNT_RNG, INTER_OPT, 0)
    TEMP_MATRIX(i, 9) = TEMP_MATRIX(i, 8) * TEMP_MATRIX(i, 6)
    TEMP_MATRIX(i, 10) = (TEMP_MATRIX(i, 7) / TEMP_MATRIX(i, 8) - 1) / TEMP_MATRIX(i, 4) + RECEIVES_BPS / 10000
    TEMP_MATRIX(i, 11) = TEMP_MATRIX(i, 3) * TEMP_MATRIX(i, 10) * TEMP_MATRIX(i, 4)
    TEMP_MATRIX(i, 12) = TEMP_MATRIX(i, 11) * TEMP_MATRIX(i, 8)
    
    TEMP1_SUM = TEMP1_SUM + TEMP_MATRIX(i, 9)
    TEMP2_SUM = TEMP2_SUM + TEMP_MATRIX(i, 12)
Next i

Select Case OUTPUT
Case 0
    SWAP_ARBITRAGE_FUNC = TEMP_MATRIX
Case Else
    ReDim TEMP_MATRIX(1 To 4, 1 To 2)
    
    TEMP_MATRIX(1, 1) = "NPV (PAY LEG) "
    TEMP_MATRIX(2, 1) = "NPV (REC LEG) "
    TEMP_MATRIX(3, 1) = "PROFIT TO THE BANK ($)"
    TEMP_MATRIX(4, 1) = "PROFIT TO THE BANK (BPS)"
    
    TEMP_MATRIX(1, 2) = TEMP1_SUM
    TEMP_MATRIX(2, 2) = TEMP2_SUM
    TEMP_MATRIX(3, 2) = TEMP_MATRIX(1, 2) + TEMP_MATRIX(2, 2)
    TEMP_MATRIX(4, 2) = (TEMP_MATRIX(3, 2) / NOTIONAL) * 10000
    SWAP_ARBITRAGE_FUNC = TEMP_MATRIX
End Select

Exit Function
ERROR_LABEL:
SWAP_ARBITRAGE_FUNC = Err.number
End Function


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : CALLABLE_DEPOSIT_ARBITRAGE_FUNC
'DESCRIPTION   : Callable Deposit Arbitrage Model
'LIBRARY       : SWAP
'GROUP         : ARBITRAGE
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'**********************************************************************************
'**********************************************************************************

Function CALLABLE_DEPOSIT_ARBITRAGE_FUNC(ByVal CALL_DATE As Date, _
ByVal SWAP_SIGMA As Double, _
ByVal NOTIONAL As Double, _
ByVal FIXED_RATE_CLIENT As Double, _
ByVal RECEIVES_BPS As Double, _
ByVal START_DATE As Date, _
ByVal END_DATE As Date, _
ByVal FREQUENCY As Integer, _
ByRef DATES_RNG As Variant, _
ByRef DISCOUNT_RNG As Variant, _
Optional ByVal COUNT_BASIS As Integer = 3, _
Optional ByVal INTER_OPT As Integer = 0)

'SWAP_SIGMA: Swap rate volatility

Dim i As Long
Dim NROWS As Long
Dim NCOLUMNS As Long

Dim SWAP_ANNUITY As Double
Dim FORWARD_ANNUITY As Double

Dim TEMP_VECTOR As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

TEMP_VECTOR = SWAP_ARBITRAGE_FUNC(NOTIONAL, FIXED_RATE_CLIENT, RECEIVES_BPS, START_DATE, END_DATE, FREQUENCY, DATES_RNG, DISCOUNT_RNG, 0, COUNT_BASIS, INTER_OPT)
'Annuity Calculation(Forward & Swap):

NROWS = UBound(TEMP_VECTOR, 1)
NCOLUMNS = UBound(TEMP_VECTOR, 2) + 2

ReDim Preserve TEMP_VECTOR(0 To NROWS, 1 To NCOLUMNS + 2)
'One column for the Dummies of the Forward Annuities
'and the other Column for the Dummies of the Swap Annuities

For i = 1 To NROWS
    If (TEMP_VECTOR(i, 1) >= CALL_DATE) And (TEMP_VECTOR(i, 2) <= END_DATE) Then
        FORWARD_ANNUITY = 1
    Else
        FORWARD_ANNUITY = 0
    End If
    If (TEMP_VECTOR(i, 2) <= END_DATE) Then
        SWAP_ANNUITY = 1
    Else
        SWAP_ANNUITY = 0
    End If
    TEMP_VECTOR(i, NCOLUMNS - 1) = FORWARD_ANNUITY
    TEMP_VECTOR(i, NCOLUMNS) = SWAP_ANNUITY
Next i

ReDim TEMP_MATRIX(1 To 14, 1 To 2)

TEMP_MATRIX(1, 1) = "CALLABLE_DEPOSIT_ARBITRAGE_FUNC"
TEMP_MATRIX(2, 1) = "OPTION MATURITY" 'SAME AS CALL DATE
TEMP_MATRIX(3, 1) = "FORWARD ANNUITY"
TEMP_MATRIX(4, 1) = "FORWARD SWAP"
TEMP_MATRIX(5, 1) = "STRIKE"
TEMP_MATRIX(6, 1) = "VOLATILITY"
TEMP_MATRIX(7, 1) = "SWAPTION PREMIUM"
TEMP_MATRIX(8, 1) = "SWAP ANNUITY"
TEMP_MATRIX(9, 1) = "ENHANCED YIELD"
TEMP_MATRIX(10, 1) = "NPV (PAY LEG)"
TEMP_MATRIX(11, 1) = "NPV (REC LEG)"
TEMP_MATRIX(12, 1) = "SWAPTION PREMIUM"
TEMP_MATRIX(13, 1) = "PROFIT TO THE BANK($)"
TEMP_MATRIX(14, 1) = "PROFIT TO THE BANK (%NOTIONAL)"

TEMP_MATRIX(1, 2) = ""
TEMP_MATRIX(2, 2) = CALL_DATE

TEMP_MATRIX(3, 2) = 0
TEMP_MATRIX(5, 2) = 0
TEMP_MATRIX(8, 2) = 0
TEMP_MATRIX(10, 2) = 0
TEMP_MATRIX(11, 2) = 0
For i = 1 To NROWS
    TEMP_MATRIX(3, 2) = TEMP_MATRIX(3, 2) + TEMP_VECTOR(i, NCOLUMNS - 1) * _
                        TEMP_VECTOR(i, 4) * TEMP_VECTOR(i, 8)

    TEMP_MATRIX(5, 2) = TEMP_MATRIX(5, 2) + TEMP_VECTOR(i, NCOLUMNS - 1) * _
                        TEMP_VECTOR(i, 4) * TEMP_VECTOR(i, 5) * TEMP_VECTOR(i, 8)

    TEMP_MATRIX(8, 2) = TEMP_MATRIX(8, 2) + TEMP_VECTOR(i, NCOLUMNS) * _
                        TEMP_VECTOR(i, 4) * TEMP_VECTOR(i, 8)
    TEMP_MATRIX(10, 2) = TEMP_MATRIX(10, 2) + TEMP_VECTOR(i, 9)
    TEMP_MATRIX(11, 2) = TEMP_MATRIX(11, 2) + TEMP_VECTOR(i, 12)
Next i

TEMP_MATRIX(4, 2) = (YIELD_INTERPOLATION_FUNC(TEMP_MATRIX(2, 2), DATES_RNG, DISCOUNT_RNG, INTER_OPT, 0) - YIELD_INTERPOLATION_FUNC(END_DATE, DATES_RNG, DISCOUNT_RNG, INTER_OPT, 0)) / TEMP_MATRIX(3, 2)
TEMP_MATRIX(5, 2) = TEMP_MATRIX(5, 2) / TEMP_MATRIX(3, 2)
TEMP_MATRIX(6, 2) = SWAP_SIGMA
TEMP_MATRIX(7, 2) = SWAP_CAP_FLOOR_FUNC(TEMP_MATRIX(4, 2), TEMP_MATRIX(5, 2), YEARFRAC_FUNC(START_DATE, TEMP_MATRIX(2, 2), COUNT_BASIS), TEMP_MATRIX(6, 2), -1, 0) * TEMP_MATRIX(3, 2)
TEMP_MATRIX(9, 2) = ((1 - YIELD_INTERPOLATION_FUNC(END_DATE, DATES_RNG, DISCOUNT_RNG, INTER_OPT, 0)) / TEMP_MATRIX(8, 2)) + (TEMP_MATRIX(7, 2) / TEMP_MATRIX(8, 2))
TEMP_MATRIX(12, 2) = TEMP_MATRIX(7, 2) * NOTIONAL
TEMP_MATRIX(13, 2) = TEMP_MATRIX(10, 2) + TEMP_MATRIX(11, 2) + TEMP_MATRIX(12, 2)
TEMP_MATRIX(14, 2) = (TEMP_MATRIX(13, 2) / NOTIONAL)

CALLABLE_DEPOSIT_ARBITRAGE_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
CALLABLE_DEPOSIT_ARBITRAGE_FUNC = Err.number
End Function

'**********************************************************************************
'**********************************************************************************
'FUNCTION      : SWAP_PROFIT_TABLE_FUNC
'DESCRIPTION   : Plain Vanilla Swap Profit Table
'LIBRARY       : SWAP
'GROUP         : ARBITRAGE
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'**********************************************************************************
'**********************************************************************************


Function SWAP_PROFIT_TABLE_FUNC(ByVal ZERO_BOND_PRICE As Double, _
ByVal NOTIONAL As Double, _
ByVal FIXED_RATE_CLIENT As Double, _
ByVal RECEIVES_BPS As Double, _
ByVal START_DATE As Date, _
ByVal END_DATE As Date, _
ByVal FREQUENCY As Integer, _
ByRef DATES_RNG As Variant, _
ByRef DISCOUNT_RNG As Variant, _
Optional ByVal COUNT_BASIS As Integer = 3, _
Optional ByVal INTER_OPT As Integer = 0)

'ZERO_BOND_PRICE Must be entered as a percentage of the face value

Dim TEMP_VALUE As Double
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

TEMP_MATRIX = SWAP_ARBITRAGE_FUNC(NOTIONAL, FIXED_RATE_CLIENT, RECEIVES_BPS, START_DATE, END_DATE, FREQUENCY, DATES_RNG, DISCOUNT_RNG, 1, COUNT_BASIS, INTER_OPT)
TEMP_VALUE = (NOTIONAL / ZERO_BOND_PRICE - NOTIONAL) * YIELD_INTERPOLATION_FUNC(END_DATE, DATES_RNG, DISCOUNT_RNG, INTER_OPT, 0)

TEMP_MATRIX(1, 2) = TEMP_VALUE
TEMP_MATRIX(3, 2) = TEMP_MATRIX(2, 2) - TEMP_MATRIX(1, 2)
TEMP_MATRIX(4, 2) = (TEMP_MATRIX(3, 2) / NOTIONAL) * 10000

SWAP_PROFIT_TABLE_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
SWAP_PROFIT_TABLE_FUNC = Err.number
End Function


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : SWAP_CAP_FLOOR_FUNC
'DESCRIPTION   : CAP FLOOR OPTION VALUATION
'LIBRARY       : FIXED_INCOME
'GROUP         : SWAP_ARBITRAGE
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'**********************************************************************************
'**********************************************************************************

Function SWAP_CAP_FLOOR_FUNC(ByVal FORWARD As Double, _
ByVal STRIKE As Double, _
ByVal EXPIRATION As Double, _
ByVal SIGMA As Double, _
Optional ByVal OPTION_FLAG As Integer = 1, _
Optional ByVal CND_TYPE As Integer = 0)

'FORWARD: Underlying swap rate
'STRIKE: STRIKE price
'EXPIRATION: Time to maturity
'SIGMA: Swap rate volatility

Dim D1_VAL As Double
Dim D2_VAL As Double
    
On Error GoTo ERROR_LABEL
    
D1_VAL = (Log(FORWARD / STRIKE) + 0.5 * SIGMA ^ 2 * EXPIRATION) / (SIGMA * Sqr(EXPIRATION))
D2_VAL = D1_VAL - SIGMA * Sqr(EXPIRATION)
    
Select Case OPTION_FLAG
Case 1 ', "Call", "c" 'Payer swaption --> Cap
    SWAP_CAP_FLOOR_FUNC = FORWARD * CND_FUNC(D1_VAL, CND_TYPE) - STRIKE * CND_FUNC(D2_VAL, CND_TYPE)
Case -1 ', "Put", "p" 'Receiver swaption --> Floor
    SWAP_CAP_FLOOR_FUNC = STRIKE * CND_FUNC(-D2_VAL, CND_TYPE) - FORWARD * CND_FUNC(-D1_VAL, CND_TYPE)
Case 2 'Digital Cap
    SWAP_CAP_FLOOR_FUNC = CND_FUNC(D2_VAL, CND_TYPE)
Case -2 'Digital Floor
    SWAP_CAP_FLOOR_FUNC = CND_FUNC(-D2_VAL, CND_TYPE)
Case Else
    GoTo ERROR_LABEL
End Select

Exit Function
ERROR_LABEL:
SWAP_CAP_FLOOR_FUNC = Err.number
End Function
