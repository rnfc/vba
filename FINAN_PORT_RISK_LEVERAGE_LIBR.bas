Attribute VB_Name = "FINAN_PORT_RISK_LEVERAGE_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_LEVERAGE_FUNC
'DESCRIPTION   : Leveraged investing of money borrowed at x.x% for t years:
'your Annual Gain (as a percentage of Amount Borrowed)
'vs.  Return on Investment.
'LIBRARY       : PORTFOLIO
'GROUP         : RISK_LEVERAGE
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function PORT_LEVERAGE_FUNC(ByRef SETTLEMENT_RNG As Variant, _
ByRef MATURITY_RNG As Variant, _
ByRef AMOUNT_BORROWED_RNG As Variant, _
ByRef BORROWING_RATE_RNG As Variant, _
ByRef MARGINAL_TAX_RATE_RNG As Variant, _
ByRef WEIGHT_TAX_RATE_RNG As Variant, _
ByRef ROI_RNG As Variant, _
Optional ByRef COUNT_BASIS_RNG As Variant = 0, _
Optional ByVal OUTPUT As Integer = 0)

Dim j As Long

Dim NCOLUMNS As Long

Dim TENOR As Double
Dim SETTLEMENT As Date
Dim MATURITY As Date
Dim AMOUNT_BORROWED As Double
Dim BORROWING_RATE As Double
Dim MARGINAL_TAX_RATE As Double
Dim WEIGHT_TAX_RATE As Double
Dim ROI As Double
Dim COUNT_BASIS As Integer

Dim SETTLEMENT_VECTOR As Variant
Dim MATURITY_VECTOR As Variant
Dim AMOUNT_BORROWED_VECTOR As Variant
Dim BORROWING_RATE_VECTOR As Variant
Dim MARGINAL_TAX_RATE_VECTOR As Variant
Dim WEIGHT_TAX_RATE_VECTOR As Variant
Dim ROI_VECTOR As Variant
Dim COUNT_BASIS_VECTOR As Variant

Dim TEMP_VECTOR As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

If IsArray(SETTLEMENT_RNG) = False Then
    ReDim SETTLEMENT_VECTOR(1 To 1, 1 To 1)
    SETTLEMENT_VECTOR(1, 1) = SETTLEMENT_RNG
Else
    SETTLEMENT_VECTOR = SETTLEMENT_RNG
    If UBound(SETTLEMENT_VECTOR, 2) = 1 Then
        SETTLEMENT_VECTOR = MATRIX_TRANSPOSE_FUNC(SETTLEMENT_VECTOR)
    End If
End If
NCOLUMNS = UBound(SETTLEMENT_VECTOR, 2)


If IsArray(MATURITY_RNG) = False Then
    ReDim MATURITY_VECTOR(1 To 1, 1 To 1)
    MATURITY_VECTOR(1, 1) = MATURITY_RNG
Else
    MATURITY_VECTOR = MATURITY_RNG
    If UBound(MATURITY_VECTOR, 2) = 1 Then
        MATURITY_VECTOR = MATRIX_TRANSPOSE_FUNC(MATURITY_VECTOR)
    End If
End If


If IsArray(AMOUNT_BORROWED_RNG) = False Then
    ReDim AMOUNT_BORROWED_VECTOR(1 To 1, 1 To 1)
    AMOUNT_BORROWED_VECTOR(1, 1) = AMOUNT_BORROWED_RNG
Else
    AMOUNT_BORROWED_VECTOR = AMOUNT_BORROWED_RNG
    If UBound(AMOUNT_BORROWED_VECTOR, 2) = 1 Then
        AMOUNT_BORROWED_VECTOR = MATRIX_TRANSPOSE_FUNC(AMOUNT_BORROWED_VECTOR)
    End If
End If


If IsArray(BORROWING_RATE_RNG) = False Then
    ReDim BORROWING_RATE_VECTOR(1 To 1, 1 To 1)
    BORROWING_RATE_VECTOR(1, 1) = BORROWING_RATE_RNG
Else
    BORROWING_RATE_VECTOR = BORROWING_RATE_RNG
    If UBound(BORROWING_RATE_VECTOR, 2) = 1 Then
        BORROWING_RATE_VECTOR = MATRIX_TRANSPOSE_FUNC(BORROWING_RATE_VECTOR)
    End If
End If


If IsArray(MARGINAL_TAX_RATE_RNG) = False Then
    ReDim MARGINAL_TAX_RATE_VECTOR(1 To 1, 1 To 1)
    MARGINAL_TAX_RATE_VECTOR(1, 1) = MARGINAL_TAX_RATE_RNG
Else
    MARGINAL_TAX_RATE_VECTOR = MARGINAL_TAX_RATE_RNG
    If UBound(MARGINAL_TAX_RATE_VECTOR, 2) = 1 Then
        MARGINAL_TAX_RATE_VECTOR = MATRIX_TRANSPOSE_FUNC(MARGINAL_TAX_RATE_VECTOR)
    End If
End If


If IsArray(WEIGHT_TAX_RATE_RNG) = False Then
    ReDim WEIGHT_TAX_RATE_VECTOR(1 To 1, 1 To 1)
    WEIGHT_TAX_RATE_VECTOR(1, 1) = WEIGHT_TAX_RATE_RNG
Else
    WEIGHT_TAX_RATE_VECTOR = WEIGHT_TAX_RATE_RNG
    If UBound(WEIGHT_TAX_RATE_VECTOR, 2) = 1 Then
        WEIGHT_TAX_RATE_VECTOR = MATRIX_TRANSPOSE_FUNC(WEIGHT_TAX_RATE_VECTOR)
    End If
End If


If IsArray(ROI_RNG) = False Then
    ReDim ROI_VECTOR(1 To 1, 1 To 1)
    ROI_VECTOR(1, 1) = ROI_RNG
Else
    ROI_VECTOR = ROI_RNG
    If UBound(ROI_VECTOR, 2) = 1 Then
        ROI_VECTOR = MATRIX_TRANSPOSE_FUNC(ROI_VECTOR)
    End If
End If


If IsArray(COUNT_BASIS_RNG) = False Then
    ReDim COUNT_BASIS_VECTOR(1 To 1, 1 To NCOLUMNS)
    For j = 1 To NCOLUMNS
        COUNT_BASIS_VECTOR(1, j) = 0
    Next j
Else
    COUNT_BASIS_VECTOR = COUNT_BASIS_RNG
    If UBound(COUNT_BASIS_VECTOR, 2) = 1 Then
        COUNT_BASIS_VECTOR = MATRIX_TRANSPOSE_FUNC(COUNT_BASIS_VECTOR)
    End If
End If

ReDim TEMP_MATRIX(1 To 10, 1 To NCOLUMNS + 1)

TEMP_MATRIX(1, 1) = "Final Portfolio"
TEMP_MATRIX(2, 1) = "Gross Gain in Portfolio"
TEMP_MATRIX(3, 1) = "Gross Gain in Portfolio, after paying taxes"
TEMP_MATRIX(4, 1) = "Accrued Interest on the Loan"
TEMP_MATRIX(5, 1) = "Tax Refund"
TEMP_MATRIX(6, 1) = "Loan Cost (Interest Cost, after taxes)"
TEMP_MATRIX(7, 1) = "Loan Annualized"
TEMP_MATRIX(9, 1) = "Net (after tax) Gain, after Borrowing Cost"
TEMP_MATRIX(8, 1) = "Loan per dollar invested"
TEMP_MATRIX(10, 1) = "As a Percentage of the Amount Borrowed"

For j = 1 To NCOLUMNS
    SETTLEMENT = SETTLEMENT_VECTOR(1, j)
    MATURITY = MATURITY_VECTOR(1, j)
    AMOUNT_BORROWED = AMOUNT_BORROWED_VECTOR(1, j)
    BORROWING_RATE = BORROWING_RATE_VECTOR(1, j)
    MARGINAL_TAX_RATE = MARGINAL_TAX_RATE_VECTOR(1, j)
    WEIGHT_TAX_RATE = WEIGHT_TAX_RATE_VECTOR(1, j)
    ROI = ROI_VECTOR(1, j)
    COUNT_BASIS = COUNT_BASIS_VECTOR(1, j)
    TENOR = YEARFRAC_FUNC(SETTLEMENT, MATURITY, COUNT_BASIS)
    
    TEMP_MATRIX(1, 1 + j) = AMOUNT_BORROWED * (1 + ROI) ^ TENOR
    
    TEMP_MATRIX(2, 1 + j) = TEMP_MATRIX(1, 1 + j) - AMOUNT_BORROWED
    
    TEMP_MATRIX(3, 1 + j) = TEMP_MATRIX(2, 1 + j) * (1 - WEIGHT_TAX_RATE * MARGINAL_TAX_RATE)
    
    TEMP_MATRIX(4, 1 + j) = BORROWING_RATE * AMOUNT_BORROWED * TENOR
    
    TEMP_MATRIX(5, 1 + j) = TEMP_MATRIX(4, 1 + j) * MARGINAL_TAX_RATE
    
    TEMP_MATRIX(6, 1 + j) = TEMP_MATRIX(4, 1 + j) * (1 - MARGINAL_TAX_RATE)
    
    TEMP_MATRIX(7, 1 + j) = TEMP_MATRIX(6, 1 + j) / TENOR
    
    TEMP_MATRIX(9, 1 + j) = TEMP_MATRIX(3, 1 + j) - TEMP_MATRIX(6, 1 + j)
    
    TEMP_MATRIX(8, 1 + j) = TEMP_MATRIX(9, 1 + j) / AMOUNT_BORROWED
    
    TEMP_MATRIX(10, 1 + j) = (1 + TEMP_MATRIX(9, 1 + j) / AMOUNT_BORROWED) ^ (1 / TENOR) - 1
Next j

If OUTPUT = 0 Then
    PORT_LEVERAGE_FUNC = TEMP_MATRIX
    Exit Function
End If

ReDim TEMP_VECTOR(1 To 8, 1 To j)

For j = 1 To NCOLUMNS

    SETTLEMENT = SETTLEMENT_VECTOR(1, j)
    MATURITY = MATURITY_VECTOR(1, j)
    AMOUNT_BORROWED = AMOUNT_BORROWED_VECTOR(1, j)
    BORROWING_RATE = BORROWING_RATE_VECTOR(1, j)
    MARGINAL_TAX_RATE = MARGINAL_TAX_RATE_VECTOR(1, j)
    WEIGHT_TAX_RATE = WEIGHT_TAX_RATE_VECTOR(1, j)
    ROI = ROI_VECTOR(1, j)
    COUNT_BASIS = COUNT_BASIS_VECTOR(1, j)
    TENOR = YEARFRAC_FUNC(SETTLEMENT, MATURITY, COUNT_BASIS)

    TEMP_VECTOR(1, j) = "You borrow " & _
    Format(AMOUNT_BORROWED, "$0,000") & " at " & _
    Format(BORROWING_RATE, "0.0%") & " for " & Format(TENOR, "0.0") & _
    " years, in order to invest it at " & Format(ROI, "0.0%")
    
    TEMP_VECTOR(2, j) = "After " & Format(TENOR, "0.0") & _
    " years, your investment has grown to " & _
    Format(TEMP_MATRIX(1, 1 + j), "$0,000") & _
    ", a gain of " & Format(TEMP_MATRIX(2, 1 + j), "$0,000") & "."
    
    TEMP_VECTOR(3, j) = "You pay the taxes (at " & _
    Format(WEIGHT_TAX_RATE, "0.0%") & _
    " of your marginal tax rate) and have " & _
    Format(TEMP_MATRIX(3, 1 + j), "$0,000") & " left."
    
    TEMP_VECTOR(4, j) = "The interest on the loan, namely " & _
    Format(BORROWING_RATE, "0.0%") & " x " & _
    Format(AMOUNT_BORROWED, "$0,000") & " x " & _
    Format(TENOR, "0.0") & " years = " & Format(TEMP_MATRIX(4, 1 + j), "$0,000")
    
    TEMP_VECTOR(5, j) = "is due, but you get a " & _
    Format(MARGINAL_TAX_RATE, "0.0%") & _
    " tax deduction. After taxes, your borrowing cost"
    
    TEMP_VECTOR(6, j) = "is just " & _
    Format(TEMP_MATRIX(6, 1 + j), "$0,000") & _
    ", leaving you (finally!) with a net gain of " & _
    Format(TEMP_MATRIX(3, 1 + j), "$0,000") & _
    "-" & Format(TEMP_MATRIX(6, 1 + j), "$0,000") & " = " & _
    Format(TEMP_MATRIX(9, 1 + j), "$0,000") & "."
    
    TEMP_VECTOR(7, j) = "As a percentage of the Amount Borrowed, that's a " & _
    Format((1 + TEMP_MATRIX(9, 1 + j) / AMOUNT_BORROWED) ^ (1 / TENOR) - _
    1, "0.0%") & " per year 'Return',"
    
    TEMP_VECTOR(8, j) = "after taxes. " & IIf(TEMP_MATRIX(10, 1 + j) > 0, _
    "Of course, it didn't cost you anything,", "… and it cost you " & _
    Format(-TEMP_MATRIX(9, 1 + j), "$0,000")) & " out-of-pocket."
    
Next j

If OUTPUT = 1 Then
    PORT_LEVERAGE_FUNC = TEMP_VECTOR
Else
    PORT_LEVERAGE_FUNC = Array(TEMP_MATRIX, TEMP_VECTOR)
End If

Exit Function
ERROR_LABEL:
PORT_LEVERAGE_FUNC = Err.number
End Function
