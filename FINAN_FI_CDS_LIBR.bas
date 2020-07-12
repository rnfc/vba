Attribute VB_Name = "FINAN_FI_CDS_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : CDSW_BOND_FAIR_VALUE_FUNC
'DESCRIPTION   : JPMorgan CDSW Bond Fair Value Calcuation
'Discounted Value of the coupons expected plus discounted
'value of the recovery rate assuming default.  The basis on a bond is the
'difference between this fair value and the actual bond dirty price,
'converted to a spread.
'LIBRARY       : FIXED_INCOME
'GROUP         : CDS
'ID            : 001
'LAST UPDATE   : 08-27-2009
'MOTIVATION    : MR. MIKE BOWICK
'REFERENCE     : Credit Derivatives A Primer - JP Morgan Credit Derivatives and
'Quantitative Research, January 2005
'************************************************************************************
'************************************************************************************

Function CDSW_BOND_FAIR_VALUE_FUNC(ByVal NOTIONAL As Double, _
ByVal INIT_SPREAD As Double, _
ByVal MARKET_SPREAD As Double, _
ByVal RECOVERY_RATE As Double, _
ByRef TENOR_RNG As Variant, _
ByRef SWAP_RNG As Variant, _
ByVal COUPON As Double, _
ByVal FREQUENCY As Integer, _
Optional ByVal OUTPUT As Integer = 0)

'INIT_SPREAD --> contract coupon
'MARKET_SPREAD --> current CDS market spread

Dim i As Long
Dim NROWS As Long

Dim TEMP1_SUM As Double
Dim TEMP2_SUM As Double

Dim BOND_VAL As Double 'BOND_FAIR_PRICE
Dim FAIR_VAL As Double 'FAIR_VALUE_PRICE

Dim TEMP_VECTOR As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

TEMP_MATRIX = CDSW_TABLE_FUNC(NOTIONAL, INIT_SPREAD, MARKET_SPREAD, RECOVERY_RATE, TENOR_RNG, SWAP_RNG, FREQUENCY, 0)
If OUTPUT = 2 Then
    CDSW_BOND_FAIR_VALUE_FUNC = TEMP_MATRIX
    Exit Function
ElseIf OUTPUT > 2 Then
    CDSW_BOND_FAIR_VALUE_FUNC = CDSW_TABLE_FUNC(NOTIONAL, INIT_SPREAD, MARKET_SPREAD, RECOVERY_RATE, TENOR_RNG, SWAP_RNG, FREQUENCY, 1)
    Exit Function
End If

NROWS = UBound(TEMP_MATRIX, 1)

ReDim TEMP_VECTOR(0 To NROWS, 1 To 3)

TEMP_VECTOR(0, 1) = "COUPON BOND CF" 'Thi column illustrate the concepts
'behind determining the fair value of a bond.  The concepts are the same as
'that used for the default swap.
TEMP_VECTOR(0, 2) = "DISC. CF"
TEMP_VECTOR(0, 3) = "DISC. RECOVER RATE"

TEMP_VECTOR(NROWS, 1) = COUPON * NOTIONAL / FREQUENCY + NOTIONAL
TEMP_VECTOR(NROWS, 2) = TEMP_VECTOR(NROWS, 1) * TEMP_MATRIX(NROWS, 6) * TEMP_MATRIX(NROWS, 8)
TEMP_VECTOR(NROWS, 3) = (1 - RECOVERY_RATE) * NOTIONAL * TEMP_MATRIX(NROWS, 6) * TEMP_MATRIX(NROWS, 7)

TEMP1_SUM = TEMP_VECTOR(NROWS, 2)
TEMP2_SUM = TEMP_VECTOR(NROWS, 3)

For i = (NROWS - FREQUENCY) To FREQUENCY Step -FREQUENCY
    TEMP_VECTOR(i, 1) = COUPON * NOTIONAL / FREQUENCY
    TEMP_VECTOR(i, 2) = TEMP_VECTOR(i, 1) * TEMP_MATRIX(i, 6) * TEMP_MATRIX(i, 8)
    TEMP_VECTOR(i, 3) = (1 - RECOVERY_RATE) * NOTIONAL * TEMP_MATRIX(i, 6) * TEMP_MATRIX(i, 7)
    
    TEMP1_SUM = TEMP1_SUM + TEMP_VECTOR(i, 2)
    TEMP2_SUM = TEMP2_SUM + TEMP_VECTOR(i, 3)
Next i

'-------------------------------------------------------------------------------------
Select Case OUTPUT
'-------------------------------------------------------------------------------------
Case 0 'This output illustrates the concepts behind determining the fair
'value of a bond.  The concepts are the same as that used for the default
'swap.  The fair value is defined as the discounted value of the coupons
'expected plus discounted value of the recovery rate assuming default.
'The basis on a bond is the difference between this fair value and the actual
'bond dirty price, converted to a spread.
'-------------------------------------------------------------------------------------
    BOND_VAL = TEMP1_SUM + TEMP2_SUM
    FAIR_VAL = BOND_VAL / NOTIONAL * 100

    CDSW_BOND_FAIR_VALUE_FUNC = Array(BOND_VAL, FAIR_VAL)
'-------------------------------------------------------------------------------------
Case Else
'-------------------------------------------------------------------------------------
    CDSW_BOND_FAIR_VALUE_FUNC = TEMP_VECTOR
'-------------------------------------------------------------------------------------
End Select
'-------------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
CDSW_BOND_FAIR_VALUE_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : CDSW_TABLE_FUNC
'DESCRIPTION   : JPMorgan CDSW Example Calculations Model
'LIBRARY       : FIXED_INCOME
'GROUP         : CDS
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 08-27-2009
'MOTIVATION    : MR. MIKE BOWICK
'REFERENCE     : Credit Derivatives A Primer - JP Morgan Credit Derivatives and
'Quantitative Research, January 2005
'************************************************************************************
'************************************************************************************

Private Function CDSW_TABLE_FUNC(ByVal NOTIONAL As Double, _
ByVal INIT_SPREAD As Double, _
ByVal MARKET_SPREAD As Double, _
ByVal RECOVERY_RATE As Double, _
ByRef TENOR_RNG As Variant, _
ByRef SWAP_RNG As Variant, _
Optional ByVal FREQUENCY As Integer = 0, _
Optional ByVal OUTPUT As Integer = 1)

Dim i As Long
Dim NROWS As Long

Dim TEMP1_SUM As Double
Dim TEMP2_SUM As Double
Dim TEMP3_SUM As Double

Dim CLEAN_VAL As Double 'CLEAN_SPREAD

Dim TEMP_MATRIX As Variant
Dim TEMP_VECTOR As Variant

Dim SWAP_VECTOR As Variant
Dim TENOR_VECTOR As Variant

'Notional Value: Enter notional position amount
'Initial contract spread: Enter the original contract coupon
'Current at market Spd: Enter the current market spread
'Recovery rate: Enter the assumed Recovery Rate

On Error GoTo ERROR_LABEL

TEMP1_SUM = 0
TEMP2_SUM = 0
TEMP3_SUM = 0

SWAP_VECTOR = SWAP_RNG
If UBound(SWAP_VECTOR, 1) = 1 Then
    SWAP_VECTOR = MATRIX_TRANSPOSE_FUNC(SWAP_VECTOR)
End If

TENOR_VECTOR = TENOR_RNG
If UBound(TENOR_VECTOR, 1) = 1 Then
    TENOR_VECTOR = MATRIX_TRANSPOSE_FUNC(TENOR_VECTOR)
End If

If UBound(SWAP_VECTOR, 1) <> UBound(TENOR_VECTOR, 1) Then: GoTo ERROR_LABEL

If FREQUENCY = 0 Then
    FREQUENCY = (1 / (TENOR_VECTOR(1, 1) - 0))
End If

NROWS = UBound(SWAP_VECTOR, 1)

CLEAN_VAL = MARKET_SPREAD / (1 - RECOVERY_RATE)
'Clean spread: Equal to current market spread / (1 - Recovery Rate).
'It is the annual default probability.  This is an approximation
'which is correct if one is doing the calculations assuming
'continuous possibility of default.  It is slightly off when
'assuming default is only possible on the quarterly payment
'dates, as we do below.


ReDim TEMP_MATRIX(0 To NROWS, 1 To 11)
TEMP_MATRIX(0, 1) = ("PERIOD")
TEMP_MATRIX(0, 2) = ("CF EARNED ON CDS")
'COLUMN 2: NOTIONAL VALUE * ORIGINAL CONTRACT SPREAD / 4  (FOR QUARTERLY PAYMENT)
TEMP_MATRIX(0, 3) = ("PREMIUM PAID ON OFFSETTING CDS")
'COLUMN 3: NOTIONAL VALUE * CURRENT MARKET SPREAD / 4 (FOR QUARTERLY PAYMENTS
TEMP_MATRIX(0, 4) = ("NET PREMIUM")
'COLUMN 4: DIFFERENCE BETWEEN COLUMN 2 AND COLUMN 3.  THIS REPRESENTS THE CASH
'FLOW WHICH IS BEING VALUED - I.E. THE VALUE OF A CDS UNWIND IS THE DISCOUNTED
'VALUE OF THIS CASH FLOW STREAM.
TEMP_MATRIX(0, 5) = ("SWAP CURVE")
'COLUMN 5: ENTER SWAP ZERO CURVE
TEMP_MATRIX(0, 6) = ("DISCOUNT FACTOR")
'COLUMN 6: DISCOUNT FACTORS BASED ON SWAP CURVE
TEMP_MATRIX(0, 7) = ("PROB OF DEFAULT")
'COLUMN 7: 1 MINUS COLUMN 8
TEMP_MATRIX(0, 8) = ("SURVIVAL PROBABILITY")
'COLUMN 8: PROBABILITY OF NO DEFAULT.
'CALCULATED AS 1/(1+ CLEAN SPREAD) ^ TIME IN YEARS.  THE CONCEPT IS THAT
'THE CLEAN SPREAD IS THE ANNUAL DEFAULT PROBABILITY.  THE PROBABILITY OF
'NOT HAVING DEFAULTED AFTER TWO YEARS (FOR EXAMPLE) IS THE PROBABILITY OF
'NO DEFAULT IN YEAR ONE TIME THE PROBABILITY OF NO DEFAULT IN YEAR TWO
TEMP_MATRIX(0, 9) = ("PV OF E[CF]")
'COLUMN 9: COLUMN 4 X COLUMN 6 X COLUMN 8. THE CASH FLOWS EXPECTED X THE
'DISCOUNT FACTOR FOR TIME X THE PROBABILITY OF RECEIVING THE CASH FLOWS.
TEMP_MATRIX(0, 10) = ("VALUE OF CURRENT COUPON")
'COLUMN 10: WHEN A CDS IS AT PAR IT MEANS THAT THE EXPECTED VALUE OF THE
'CASH FLOWS TO BE RECEIVED (WITH NO DEFAULT) IS EQUAL TO THE EXPECTED VALUE
'OF THE RECOVERY VALUE TO BE RECEIVED (IN CASE OF DEFAULT).  THESE
'CALCULATIONS ARE SHOWN IN COLUMNS 10 AND 11 BELOW.  THE COLUMNS GENERALLY
'WILL BE A LITTLE OFF BECAUSE THE CLEAN SPREAD WHICH IS CALCULATED
'AS CURRENT MARKET SPREAD / (1 - RECOVERY) IS NOT QUITE RIGHT
TEMP_MATRIX(0, 11) = ("VALUE OF LOSS IF DEFAULT")

TEMP_MATRIX(1, 1) = TENOR_VECTOR(1, 1)
TEMP_MATRIX(1, 2) = INIT_SPREAD * NOTIONAL * TEMP_MATRIX(1, 1)
TEMP_MATRIX(1, 3) = MARKET_SPREAD * NOTIONAL * TEMP_MATRIX(1, 1)
TEMP_MATRIX(1, 4) = TEMP_MATRIX(1, 2) - TEMP_MATRIX(1, 3)
TEMP_MATRIX(1, 5) = SWAP_VECTOR(1, 1)


TEMP_MATRIX(1, 6) = 1 / (1 + TEMP_MATRIX(1, 5) / FREQUENCY) ^ 1

TEMP_MATRIX(1, 8) = (1 + CLEAN_VAL) ^ (-TEMP_MATRIX(1, 1))
TEMP_MATRIX(1, 7) = 1 - TEMP_MATRIX(1, 8)


TEMP_MATRIX(1, 9) = TEMP_MATRIX(1, 4) * TEMP_MATRIX(1, 6) * TEMP_MATRIX(1, 8)
TEMP_MATRIX(1, 10) = TEMP_MATRIX(1, 3) * TEMP_MATRIX(1, 6) * TEMP_MATRIX(1, 8)

TEMP_MATRIX(1, 11) = (1 - RECOVERY_RATE) * TEMP_MATRIX(1, 6) * _
                TEMP_MATRIX(1, 7) * NOTIONAL

TEMP1_SUM = TEMP_MATRIX(1, 9)
TEMP2_SUM = TEMP_MATRIX(1, 10)
TEMP3_SUM = TEMP_MATRIX(1, 11)


For i = 2 To NROWS
    TEMP_MATRIX(i, 1) = TENOR_VECTOR(i, 1)
    TEMP_MATRIX(i, 2) = TEMP_MATRIX(i - 1, 2)
    TEMP_MATRIX(i, 3) = TEMP_MATRIX(i - 1, 3)
    TEMP_MATRIX(i, 4) = TEMP_MATRIX(i, 2) - TEMP_MATRIX(i, 3)
    TEMP_MATRIX(i, 5) = SWAP_VECTOR(i, 1)
    
    TEMP_MATRIX(i, 6) = 1 / (1 + TEMP_MATRIX(i, 5) / FREQUENCY) ^ i
    TEMP_MATRIX(i, 8) = (1 + CLEAN_VAL) ^ (-TEMP_MATRIX(i, 1))
    TEMP_MATRIX(i, 7) = TEMP_MATRIX(i - 1, 8) - TEMP_MATRIX(i, 8)
    
    TEMP_MATRIX(i, 9) = TEMP_MATRIX(i, 4) * TEMP_MATRIX(i, 6) * TEMP_MATRIX(i, 8)
    TEMP_MATRIX(i, 10) = TEMP_MATRIX(i, 3) * TEMP_MATRIX(i, 6) * TEMP_MATRIX(i, 8)
    TEMP_MATRIX(i, 11) = (1 - RECOVERY_RATE) * TEMP_MATRIX(i, 6) * _
                TEMP_MATRIX(i, 7) * NOTIONAL

    TEMP1_SUM = TEMP1_SUM + TEMP_MATRIX(i, 9)
    TEMP2_SUM = TEMP2_SUM + TEMP_MATRIX(i, 10)
    TEMP3_SUM = TEMP3_SUM + TEMP_MATRIX(i, 11)

Next i

'------------------------------------------------------------------------
Select Case OUTPUT
'------------------------------------------------------------------------
Case 0
'------------------------------------------------------------------------
    CDSW_TABLE_FUNC = TEMP_MATRIX
'------------------------------------------------------------------------
Case Else 'Mark to market of CDS unwind
'------------------------------------------------------------------------
    ReDim TEMP_VECTOR(1 To 4, 1 To 2)
    
'Mark-to-Market Value of Unwinding the CDS contract. Can be compared to CDSW.
    TEMP_VECTOR(1, 1) = ("SUM PV OF E[CF]")
    TEMP_VECTOR(2, 1) = ("SUM OF CURRENT COUPON")
    TEMP_VECTOR(3, 1) = ("SUM OF LOSS IF DEFAULT")
    TEMP_VECTOR(4, 1) = ("DIFF")
    
    TEMP_VECTOR(1, 2) = TEMP1_SUM
    TEMP_VECTOR(2, 2) = TEMP2_SUM
    TEMP_VECTOR(3, 2) = TEMP3_SUM
    TEMP_VECTOR(4, 2) = TEMP3_SUM - TEMP2_SUM
    
    CDSW_TABLE_FUNC = TEMP_VECTOR
'------------------------------------------------------------------------
End Select
'------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
CDSW_TABLE_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CDSW_BOND_FAIR_VALUE_FUNC
'DESCRIPTION   : Credit default swaps are par instruments whereas most bonds trade
'at a discount or a premium. The following routine calculates an equivalent spread
'for a corporate bond so that like comparisons can be made with credit default swaps.
'LIBRARY       : FIXED_INCOME
'GROUP         : CDS
'ID            : 003
'LAST UPDATE   : 08-27-2009
'MOTIVATION    : MR. MIKE BOWICK
'REFERENCE     : Credit Derivatives A Primer - JP Morgan Credit Derivatives and
'Quantitative Research, January 2005
'************************************************************************************
'************************************************************************************

Function CDSW_BOND_COMPARISON_FUNC(ByVal BOND_PRICE_RNG As Variant, _
ByVal SETTLEMENT_RNG As Variant, _
ByVal MATURITY_RNG As Variant, _
ByVal SWAP_RATE_MATURITY_RNG As Variant, _
ByVal COUPON_RNG As Variant, _
ByVal RECOVERY_RNG As Variant, _
Optional ByVal FREQUENCY_RNG As Variant = 2, _
Optional ByVal REDEMPTION_RNG As Variant = 100, _
Optional ByVal COUNT_BASIS_RNG As Variant = 0, _
Optional ByVal GUESS_YIELD_RNG As Variant = 0.1)

'SWAP_RATE_MATURITY --> Swap rate to bond maturity
'RECOVERY --> Estimated Recovery Rate

Dim i As Long
Dim NROWS As Long

Dim BOND_PRICE_VECTOR As Variant
Dim SETTLEMENT_VECTOR As Variant
Dim MATURITY_VECTOR As Variant
Dim SWAP_RATE_MATURITY_VECTOR As Variant

Dim COUPON_VECTOR As Variant
Dim RECOVERY_VECTOR As Variant
Dim FREQUENCY_VECTOR As Variant
Dim REDEMPTION_VECTOR As Variant
Dim COUNT_BASIS_VECTOR As Variant
Dim GUESS_YIELD_VECTOR As Variant

Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

'--------------------------------------------------------------------------------------
If IsArray(BOND_PRICE_RNG) = True Then
    BOND_PRICE_VECTOR = BOND_PRICE_RNG
    If UBound(BOND_PRICE_VECTOR, 1) = 1 Then
        BOND_PRICE_VECTOR = MATRIX_TRANSPOSE_FUNC(BOND_PRICE_RNG)
    End If
Else
    ReDim BOND_PRICE_VECTOR(1 To 1, 1 To 1)
    BOND_PRICE_VECTOR(1, 1) = BOND_PRICE_RNG
End If
NROWS = UBound(BOND_PRICE_VECTOR, 1)

'--------------------------------------------------------------------------------------
If IsArray(SETTLEMENT_RNG) = True Then
    SETTLEMENT_VECTOR = SETTLEMENT_RNG
    If UBound(SETTLEMENT_VECTOR, 1) = 1 Then
        SETTLEMENT_VECTOR = MATRIX_TRANSPOSE_FUNC(SETTLEMENT_RNG)
    End If
Else
    ReDim SETTLEMENT_VECTOR(1 To 1, 1 To 1)
    SETTLEMENT_VECTOR(1, 1) = SETTLEMENT_RNG
End If
If UBound(SETTLEMENT_VECTOR, 1) <> NROWS Then: GoTo ERROR_LABEL
'--------------------------------------------------------------------------------------
If IsArray(MATURITY_RNG) = True Then
    MATURITY_VECTOR = MATURITY_RNG
    If UBound(MATURITY_VECTOR, 1) = 1 Then
        MATURITY_VECTOR = MATRIX_TRANSPOSE_FUNC(MATURITY_RNG)
    End If
Else
    ReDim MATURITY_VECTOR(1 To 1, 1 To 1)
    MATURITY_VECTOR(1, 1) = MATURITY_RNG
End If
If UBound(MATURITY_VECTOR, 1) <> NROWS Then: GoTo ERROR_LABEL
'--------------------------------------------------------------------------------------
If IsArray(SWAP_RATE_MATURITY_RNG) = True Then
    SWAP_RATE_MATURITY_VECTOR = SWAP_RATE_MATURITY_RNG
    If UBound(MATURITY_VECTOR, 1) = 1 Then
        SWAP_RATE_MATURITY_VECTOR = MATRIX_TRANSPOSE_FUNC(SWAP_RATE_MATURITY_RNG)
    End If
Else
    ReDim SWAP_RATE_MATURITY_VECTOR(1 To 1, 1 To 1)
    SWAP_RATE_MATURITY_VECTOR(1, 1) = SWAP_RATE_MATURITY_RNG
End If
If UBound(SWAP_RATE_MATURITY_VECTOR, 1) <> NROWS Then: GoTo ERROR_LABEL
'--------------------------------------------------------------------------------------
If IsArray(COUPON_RNG) = True Then
    COUPON_VECTOR = COUPON_RNG
    If UBound(COUPON_VECTOR, 1) = 1 Then
        COUPON_VECTOR = MATRIX_TRANSPOSE_FUNC(COUPON_RNG)
    End If
Else
    ReDim COUPON_VECTOR(1 To 1, 1 To 1)
    COUPON_VECTOR(1, 1) = COUPON_RNG
End If
If UBound(COUPON_VECTOR, 1) <> NROWS Then: GoTo ERROR_LABEL
'--------------------------------------------------------------------------------------
If IsArray(RECOVERY_RNG) = True Then
    RECOVERY_VECTOR = RECOVERY_RNG
    If UBound(RECOVERY_VECTOR, 1) = 1 Then
        RECOVERY_VECTOR = MATRIX_TRANSPOSE_FUNC(RECOVERY_RNG)
    End If
Else
    ReDim RECOVERY_VECTOR(1 To 1, 1 To 1)
    RECOVERY_VECTOR(1, 1) = RECOVERY_RNG
End If
If UBound(RECOVERY_VECTOR, 1) <> NROWS Then: GoTo ERROR_LABEL
'--------------------------------------------------------------------------------------
If IsArray(FREQUENCY_RNG) = True Then
    FREQUENCY_VECTOR = FREQUENCY_RNG
    If UBound(FREQUENCY_VECTOR, 1) = 1 Then
        FREQUENCY_VECTOR = MATRIX_TRANSPOSE_FUNC(FREQUENCY_RNG)
    End If
Else
    ReDim FREQUENCY_VECTOR(1 To 1, 1 To 1)
    FREQUENCY_VECTOR(1, 1) = FREQUENCY_RNG
End If
If UBound(FREQUENCY_VECTOR, 1) <> NROWS Then: GoTo ERROR_LABEL
'--------------------------------------------------------------------------------------
If IsArray(REDEMPTION_RNG) = True Then
    REDEMPTION_VECTOR = REDEMPTION_RNG
    If UBound(REDEMPTION_VECTOR, 1) = 1 Then
        REDEMPTION_VECTOR = MATRIX_TRANSPOSE_FUNC(REDEMPTION_RNG)
    End If
Else
    ReDim REDEMPTION_VECTOR(1 To 1, 1 To 1)
    REDEMPTION_VECTOR(1, 1) = REDEMPTION_RNG
End If
If UBound(REDEMPTION_VECTOR, 1) <> NROWS Then: GoTo ERROR_LABEL
'--------------------------------------------------------------------------------------
If IsArray(COUNT_BASIS_RNG) = True Then
    COUNT_BASIS_VECTOR = COUNT_BASIS_RNG
    If UBound(COUNT_BASIS_VECTOR, 1) = 1 Then
        COUNT_BASIS_VECTOR = MATRIX_TRANSPOSE_FUNC(COUNT_BASIS_RNG)
    End If
Else
    ReDim COUNT_BASIS_VECTOR(1 To 1, 1 To 1)
    COUNT_BASIS_VECTOR(1, 1) = COUNT_BASIS_RNG
End If
If UBound(COUNT_BASIS_VECTOR, 1) <> NROWS Then: GoTo ERROR_LABEL
'--------------------------------------------------------------------------------------
If IsArray(GUESS_YIELD_RNG) = True Then
    GUESS_YIELD_VECTOR = GUESS_YIELD_RNG
    If UBound(GUESS_YIELD_VECTOR, 1) = 1 Then
        GUESS_YIELD_VECTOR = MATRIX_TRANSPOSE_FUNC(GUESS_YIELD_RNG)
    End If
Else
    ReDim GUESS_YIELD_VECTOR(1 To NROWS, 1 To 1)
    For i = 1 To NROWS
        GUESS_YIELD_VECTOR(i, 1) = GUESS_YIELD_RNG
    Next i
End If
If UBound(GUESS_YIELD_VECTOR, 1) <> NROWS Then: GoTo ERROR_LABEL
'--------------------------------------------------------------------------------------

ReDim TEMP_MATRIX(0 To NROWS, 1 To 7)

TEMP_MATRIX(0, 1) = "ZERO_RECOVERY_RATE"
TEMP_MATRIX(0, 2) = "IMPLIED_PRICE" 'Calculate the implied price and coupon of the zero recovery bond
TEMP_MATRIX(0, 3) = "COUPON" 'coupon zero recovery bond
TEMP_MATRIX(0, 4) = "SWAP SPREAD"
TEMP_MATRIX(0, 5) = "YTM"
'Calculate the yield-to-maturity for the zero recovery bond
TEMP_MATRIX(0, 6) = "CLEAN SPREAD"
'The clean spread is the return paid to investors for assuming the risk the issuer defaults

TEMP_MATRIX(0, 7) = "PAR SPREAD"
'The par spread is the spread above Libor for a given bond assuming a $100 price.


For i = 1 To NROWS
    TEMP_MATRIX(i, 1) = RECOVERY_VECTOR(i, 1)
    
    TEMP_MATRIX(i, 2) = (BOND_PRICE_VECTOR(i, 1) - (TEMP_MATRIX(i, 1) * 100)) / (1 - TEMP_MATRIX(i, 1))
    TEMP_MATRIX(i, 3) = (COUPON_VECTOR(i, 1) - (TEMP_MATRIX(i, 1) * SWAP_RATE_MATURITY_VECTOR(i, 1))) / (1 - TEMP_MATRIX(i, 1))
    TEMP_MATRIX(i, 4) = TEMP_MATRIX(i, 3) - SWAP_RATE_MATURITY_VECTOR(i, 1)
    TEMP_MATRIX(i, 5) = BOND_YIELD_FUNC(TEMP_MATRIX(i, 2), SETTLEMENT_VECTOR(i, 1), MATURITY_VECTOR(i, 1), TEMP_MATRIX(i, 3), FREQUENCY_VECTOR(i, 1), REDEMPTION_VECTOR(i, 1), COUNT_BASIS_VECTOR(i, 1), GUESS_YIELD_VECTOR(i, 1))
    'convert the YTM and swap rate from semi-annual 30/360 to continuous 30/360 rates
    TEMP_MATRIX(i, 6) = (FREQUENCY_VECTOR(i, 1) * Log(1 + TEMP_MATRIX(i, 5) / FREQUENCY_VECTOR(i, 1)))
    TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 6) - (2 * (Log(1 + SWAP_RATE_MATURITY_VECTOR(i, 1) / 2)))
    TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 6) * 360 / 365
    TEMP_MATRIX(i, 7) = TEMP_MATRIX(i, 6) * (1 - TEMP_MATRIX(i, 1))
Next i

CDSW_BOND_COMPARISON_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
CDSW_BOND_COMPARISON_FUNC = Err.number
End Function
