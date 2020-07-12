Attribute VB_Name = "FINAN_DERIV_WARRANTS_LIBR"

'--------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'--------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------

Private PUB_STOCK_PRICE_VAL As Double
Private PUB_STRIKE_VAL As Double
Private PUB_EXPIRATION_VAL As Double
Private PUB_RATE_VAL As Double
Private PUB_YIELD_VAL As Double
Private PUB_VOLATILITY_VAL As Double
Private PUB_EQUITY_VAL As Double
Private PUB_SHARES_OUTSTANDING_VAL As Double
Private PUB_WARRANTS_OUTSTANDING_VAL As Double
Private PUB_ESTIMATION_FLAG As Boolean
Private PUB_OPTION_VAL As Double
Private PUB_OPTION_MODEL_VAL As Integer

Private Const PUB_EPSILON As Double = 2 ^ 52


'************************************************************************************
'************************************************************************************
'FUNCTION      : MGMT_OPTIONS_VALUATION_FUNC

'DESCRIPTION   : Valuing Management Options or Warrants when there is dilution
'This program is designed to value options, the exercise of which can
'create more shares and thus affect the stock price. This is the case
'with warrants and management options. It is also the case with convertible
'bonds. As a general rule, using an unadjusted option pricing model to value
'these options will overstate their value.

'http://www.sec.gov/answers/empopt.htm
'http://www.taxqueries.com/questions/33
'http://www.apra.gov.au/Careers/upload/Tristan-Boyd-What-Determines-Early-Exercise-of-Employee-Stock-Options.pdf

'LIBRARY       : DERIVATIVES
'GROUP         : WARRANTS
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 05/10/2011
'************************************************************************************
'************************************************************************************

Function MGMT_OPTIONS_VALUATION_FUNC( _
ByRef EQUITY_VALUE_RNG As Variant, _
ByRef SHARES_OUTSTANDING_RNG As Variant, _
ByRef WARRANTS_OUTSTANDING_RNG As Variant, _
ByRef STRIKE_PRICE_RNG As Variant, _
ByRef EXPIRATION_RNG As Variant, _
ByRef VOLATILITY_RNG As Variant, _
ByRef STOCK_PRICE_RNG As Variant, _
ByRef DIVIDEND_YIELD_RNG As Variant, _
ByRef RISK_FREE_RNG As Variant, _
Optional ByRef ESTIMATION_FLAG_RNG As Variant = False, _
Optional ByRef OPTION_MODEL_RNG As Variant = 0)

'EQUITY_VALUE: This presumably comes from a DCF model
'or by applying a multiple to earningss or book value

'STOCK_PRICE: Enter the current stock price. If you are
'assessing the value of the stock, enter your estimate.

'STRIKE: Enter the exercise price of the warrants. If you
'are valuing a portfolio of options, enter the weighted average price

'EXPIRATION: Time (in years) until expiration of the warrant.
'Enter months as a fraction of a year.

'VOLATILITY: Enter the annualized standard deviation in the stock price.
'You can use either a historical estimate or an implied standard deviation.

'DIVIDEND_YIELD: Annualized dividend yield on stock. Divide the dollar
'dividends (annual) by the stock price. You can use industry average
'standard deviations.

'RISK_FREE: Treasury bond rate. Enter the rate on a government bond with EXPIRATION
'closest to option expiration.

'WARRANTS_OUTSTANDING: Enter the number of warrants (options)
'outstanding. Enter the number of shares outstanding

'SHARES_OUTSTANDING: Primary shares outstanding currently.
'(Do not count shares that will be created by warrant exercise)


Dim i As Long
Dim j As Long
Dim k As Long
Dim NROWS As Long
Dim NCOLUMNS As Long
Dim HEADINGS_STR As String

Dim X_VAL As Double

Dim EQUITY_VALUE_VECTOR As Variant 'Aggregate value of equity: This presumably comes from
'a DCF model or by applying a multiple to earningss or book value

Dim SHARES_OUTSTANDING_VECTOR As Variant 'Primary number of shares outstanding
Dim WARRANTS_OUTSTANDING_VECTOR As Variant 'Number of options outstanding
Dim STRIKE_PRICE_VECTOR As Variant
Dim EXPIRATION_VECTOR As Variant
Dim VOLATILITY_VECTOR As Variant
Dim STOCK_PRICE_VECTOR As Variant
Dim DIVIDEND_YIELD_VECTOR As Variant
Dim RISK_FREE_VECTOR As Variant

Dim ESTIMATION_FLAG_VECTOR As Variant 'if FALSE we use the current stock price to
'estimate option value, rather than the estimated value per share.

Dim OPTION_MODEL_VECTOR As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

If IsArray(EQUITY_VALUE_RNG) = True Then
    EQUITY_VALUE_VECTOR = EQUITY_VALUE_RNG
    If UBound(EQUITY_VALUE_VECTOR, 1) = 1 Then
        EQUITY_VALUE_VECTOR = MATRIX_TRANSPOSE_FUNC(EQUITY_VALUE_VECTOR)
    End If
Else
    ReDim EQUITY_VALUE_VECTOR(1 To 1, 1 To 1)
    EQUITY_VALUE_VECTOR(1, 1) = EQUITY_VALUE_RNG
End If
NROWS = UBound(EQUITY_VALUE_VECTOR, 1)

If IsArray(SHARES_OUTSTANDING_RNG) = True Then
    SHARES_OUTSTANDING_VECTOR = SHARES_OUTSTANDING_RNG
    If UBound(SHARES_OUTSTANDING_VECTOR, 1) = 1 Then
        SHARES_OUTSTANDING_VECTOR = MATRIX_TRANSPOSE_FUNC(SHARES_OUTSTANDING_VECTOR)
    End If
Else
    ReDim SHARES_OUTSTANDING_VECTOR(1 To 1, 1 To 1)
    SHARES_OUTSTANDING_VECTOR(1, 1) = SHARES_OUTSTANDING_RNG
End If
If NROWS <> UBound(SHARES_OUTSTANDING_VECTOR, 1) Then: GoTo ERROR_LABEL

If IsArray(WARRANTS_OUTSTANDING_RNG) = True Then
    WARRANTS_OUTSTANDING_VECTOR = WARRANTS_OUTSTANDING_RNG
    If UBound(WARRANTS_OUTSTANDING_VECTOR, 1) = 1 Then
        WARRANTS_OUTSTANDING_VECTOR = MATRIX_TRANSPOSE_FUNC(WARRANTS_OUTSTANDING_VECTOR)
    End If
Else
    ReDim WARRANTS_OUTSTANDING_VECTOR(1 To 1, 1 To 1)
    WARRANTS_OUTSTANDING_VECTOR(1, 1) = WARRANTS_OUTSTANDING_RNG
End If
If NROWS <> UBound(WARRANTS_OUTSTANDING_VECTOR, 1) Then: GoTo ERROR_LABEL

If IsArray(STRIKE_PRICE_RNG) = True Then
    STRIKE_PRICE_VECTOR = STRIKE_PRICE_RNG
    If UBound(STRIKE_PRICE_VECTOR, 1) = 1 Then
        STRIKE_PRICE_VECTOR = MATRIX_TRANSPOSE_FUNC(STRIKE_PRICE_VECTOR)
    End If
Else
    ReDim STRIKE_PRICE_VECTOR(1 To 1, 1 To 1)
    STRIKE_PRICE_VECTOR(1, 1) = STRIKE_PRICE_RNG
End If
If NROWS <> UBound(STRIKE_PRICE_VECTOR, 1) Then: GoTo ERROR_LABEL

If IsArray(EXPIRATION_RNG) = True Then
    EXPIRATION_VECTOR = EXPIRATION_RNG
    If UBound(EXPIRATION_VECTOR, 1) = 1 Then
        EXPIRATION_VECTOR = MATRIX_TRANSPOSE_FUNC(EXPIRATION_VECTOR)
    End If
Else
    ReDim EXPIRATION_VECTOR(1 To 1, 1 To 1)
    EXPIRATION_VECTOR(1, 1) = EXPIRATION_RNG
End If
If NROWS <> UBound(EXPIRATION_VECTOR, 1) Then: GoTo ERROR_LABEL

If IsArray(VOLATILITY_RNG) = True Then
    VOLATILITY_VECTOR = VOLATILITY_RNG
    If UBound(VOLATILITY_VECTOR, 1) = 1 Then
        VOLATILITY_VECTOR = MATRIX_TRANSPOSE_FUNC(VOLATILITY_VECTOR)
    End If
Else
    ReDim VOLATILITY_VECTOR(1 To 1, 1 To 1)
    VOLATILITY_VECTOR(1, 1) = VOLATILITY_RNG
End If
If NROWS <> UBound(VOLATILITY_VECTOR, 1) Then: GoTo ERROR_LABEL

If IsArray(STOCK_PRICE_RNG) = True Then
    STOCK_PRICE_VECTOR = STOCK_PRICE_RNG
    If UBound(STOCK_PRICE_VECTOR, 1) = 1 Then
        STOCK_PRICE_VECTOR = MATRIX_TRANSPOSE_FUNC(STOCK_PRICE_VECTOR)
    End If
Else
    ReDim STOCK_PRICE_VECTOR(1 To 1, 1 To 1)
    STOCK_PRICE_VECTOR(1, 1) = STOCK_PRICE_RNG
End If
If NROWS <> UBound(STOCK_PRICE_VECTOR, 1) Then: GoTo ERROR_LABEL

If IsArray(DIVIDEND_YIELD_RNG) = True Then
    DIVIDEND_YIELD_VECTOR = DIVIDEND_YIELD_RNG
    If UBound(DIVIDEND_YIELD_VECTOR, 1) = 1 Then
        DIVIDEND_YIELD_VECTOR = MATRIX_TRANSPOSE_FUNC(DIVIDEND_YIELD_VECTOR)
    End If
Else
    ReDim DIVIDEND_YIELD_VECTOR(1 To 1, 1 To 1)
    DIVIDEND_YIELD_VECTOR(1, 1) = DIVIDEND_YIELD_RNG
End If
If NROWS <> UBound(DIVIDEND_YIELD_VECTOR, 1) Then: GoTo ERROR_LABEL

If IsArray(RISK_FREE_RNG) = True Then
    RISK_FREE_VECTOR = RISK_FREE_RNG
    If UBound(RISK_FREE_VECTOR, 1) = 1 Then
        RISK_FREE_VECTOR = MATRIX_TRANSPOSE_FUNC(RISK_FREE_VECTOR)
    End If
Else
    ReDim RISK_FREE_VECTOR(1 To 1, 1 To 1)
    RISK_FREE_VECTOR(1, 1) = RISK_FREE_RNG
End If
If NROWS <> UBound(RISK_FREE_VECTOR, 1) Then: GoTo ERROR_LABEL

If IsArray(ESTIMATION_FLAG_RNG) = True Then
    ESTIMATION_FLAG_VECTOR = ESTIMATION_FLAG_RNG
    If UBound(ESTIMATION_FLAG_VECTOR, 1) = 1 Then
        ESTIMATION_FLAG_VECTOR = MATRIX_TRANSPOSE_FUNC(ESTIMATION_FLAG_VECTOR)
    End If
Else
    ReDim ESTIMATION_FLAG_VECTOR(1 To NROWS, 1 To 1)
    For i = 1 To NROWS: ESTIMATION_FLAG_VECTOR(i, 1) = ESTIMATION_FLAG_RNG: Next i
End If
If NROWS <> UBound(ESTIMATION_FLAG_VECTOR, 1) Then: GoTo ERROR_LABEL

If IsArray(OPTION_MODEL_RNG) = True Then
    OPTION_MODEL_VECTOR = OPTION_MODEL_RNG
    If UBound(OPTION_MODEL_VECTOR, 1) = 1 Then
        OPTION_MODEL_VECTOR = MATRIX_TRANSPOSE_FUNC(OPTION_MODEL_VECTOR)
    End If
Else
    ReDim OPTION_MODEL_VECTOR(1 To NROWS, 1 To 1)
    For i = 1 To NROWS: OPTION_MODEL_VECTOR(i, 1) = OPTION_MODEL_RNG: Next i
End If
If NROWS <> UBound(OPTION_MODEL_VECTOR, 1) Then: GoTo ERROR_LABEL

NCOLUMNS = 16
ReDim TEMP_MATRIX(0 To NROWS, 1 To NCOLUMNS)
HEADINGS_STR = "Option Pricing Model,Value of the call,Aggregate value of equity, Value of management options,Value of equity in common stock,Primary number of shares, Value per share,Market price,% under or over valued,Treasury stock approach: Aggregate value of equity,Treasury stock approach:  + Proceeds from exercise,Treasury stock approach: / Diluted number of shares,Treasury stock approach: Value per share,Diluted shares approach: Aggregate value of equity,Diluted shares approach: / Diluted number of shares,Diluted shares approach: Value per share,"
i = 1
For k = 1 To NCOLUMNS
    j = InStr(i, HEADINGS_STR, ",")
    TEMP_MATRIX(0, k) = Mid(HEADINGS_STR, i, j - i)
    i = j + 1
Next k

For i = 1 To NROWS
    GoSub OPTIMIZER_LINE
    TEMP_MATRIX(i, 1) = OPTION_MODEL_VECTOR(i, 1) + 1
    '(1) Black Scholes
    '(2) Bjerksund and Stensland American Approximation (2002)
    '(3) Barone-Adesi and Whaley (1987) American Approximation
    TEMP_MATRIX(i, 2) = PUB_OPTION_VAL
    
    TEMP_MATRIX(i, 3) = EQUITY_VALUE_VECTOR(i, 1)
    TEMP_MATRIX(i, 4) = TEMP_MATRIX(i, 2) * WARRANTS_OUTSTANDING_VECTOR(i, 1)
    TEMP_MATRIX(i, 5) = TEMP_MATRIX(i, 3) - TEMP_MATRIX(i, 4)
    TEMP_MATRIX(i, 6) = SHARES_OUTSTANDING_VECTOR(i, 1)
    TEMP_MATRIX(i, 7) = TEMP_MATRIX(i, 5) / TEMP_MATRIX(i, 6)
    TEMP_MATRIX(i, 8) = STOCK_PRICE_VECTOR(i, 1)
    TEMP_MATRIX(i, 9) = TEMP_MATRIX(i, 8) / TEMP_MATRIX(i, 7) - 1
    
    TEMP_MATRIX(i, 10) = TEMP_MATRIX(i, 3)
    TEMP_MATRIX(i, 11) = STRIKE_PRICE_VECTOR(i, 1) * WARRANTS_OUTSTANDING_VECTOR(i, 1)
    TEMP_MATRIX(i, 12) = TEMP_MATRIX(i, 6) + WARRANTS_OUTSTANDING_VECTOR(i, 1)
    TEMP_MATRIX(i, 13) = (TEMP_MATRIX(i, 10) + TEMP_MATRIX(i, 11)) / TEMP_MATRIX(i, 12)
    
    TEMP_MATRIX(i, 14) = TEMP_MATRIX(i, 3)
    TEMP_MATRIX(i, 15) = TEMP_MATRIX(i, 12)
    TEMP_MATRIX(i, 16) = TEMP_MATRIX(i, 14) / TEMP_MATRIX(i, 15)
Next i

MGMT_OPTIONS_VALUATION_FUNC = TEMP_MATRIX

Exit Function
'---------------------------------------------------------------------------------------------------------------------------
OPTIMIZER_LINE:
'---------------------------------------------------------------------------------------------------------------------------
    PUB_ESTIMATION_FLAG = ESTIMATION_FLAG_VECTOR(i, 1)
    PUB_STOCK_PRICE_VAL = STOCK_PRICE_VECTOR(i, 1)
    PUB_STRIKE_VAL = STRIKE_PRICE_VECTOR(i, 1)
    PUB_EXPIRATION_VAL = EXPIRATION_VECTOR(i, 1)
    PUB_RATE_VAL = RISK_FREE_VECTOR(i, 1)
    PUB_YIELD_VAL = DIVIDEND_YIELD_VECTOR(i, 1)
    PUB_VOLATILITY_VAL = VOLATILITY_VECTOR(i, 1)
    PUB_EQUITY_VAL = EQUITY_VALUE_VECTOR(i, 1)
    PUB_SHARES_OUTSTANDING_VAL = SHARES_OUTSTANDING_VECTOR(i, 1)
    PUB_WARRANTS_OUTSTANDING_VAL = WARRANTS_OUTSTANDING_VECTOR(i, 1)
    PUB_OPTION_MODEL_VAL = OPTION_MODEL_VECTOR(i, 1)
    X_VAL = MULLER_ZERO_FUNC(10 ^ -10, 10 ^ 5, "MGMT_OPTIONS_OBJ_FUNC", , , 500, 0.00000000000001)
    If X_VAL = PUB_EPSILON Then: X_VAL = 0
'---------------------------------------------------------------------------------------------------------------------------
Return
'---------------------------------------------------------------------------------------------------------------------------
ERROR_LABEL:
MGMT_OPTIONS_VALUATION_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MGMT_OPTIONS_OBJ_FUNC
'DESCRIPTION   :
'LIBRARY       : DERIVATIVES
'GROUP         : WARRANTS
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 05/10/2011
'************************************************************************************
'************************************************************************************

Private Function MGMT_OPTIONS_OBJ_FUNC(ByVal X_VAL As Double)

Dim N_VAL As Double 'Diluted number of shares
Dim S_VAL As Double 'Stock Price
Dim AS_VAL As Double 'Adjusted S

On Error GoTo ERROR_LABEL

N_VAL = PUB_SHARES_OUTSTANDING_VAL + PUB_WARRANTS_OUTSTANDING_VAL
If PUB_ESTIMATION_FLAG = True Then '(Aggregate value of equity -(Value of management options))/ Primary number of shares
    S_VAL = (PUB_EQUITY_VAL - (X_VAL * PUB_WARRANTS_OUTSTANDING_VAL)) / PUB_SHARES_OUTSTANDING_VAL
Else
    S_VAL = PUB_STOCK_PRICE_VAL
End If
AS_VAL = (S_VAL * PUB_SHARES_OUTSTANDING_VAL + X_VAL * PUB_WARRANTS_OUTSTANDING_VAL) / N_VAL

Select Case PUB_OPTION_MODEL_VAL
Case 0 'Black Scholes
    PUB_OPTION_VAL = EUROPEAN_CALL_OPTION_FUNC(AS_VAL, PUB_STRIKE_VAL, PUB_EXPIRATION_VAL, PUB_RATE_VAL, PUB_YIELD_VAL, PUB_VOLATILITY_VAL, 0)
Case 1 'Bjerksund and Stensland American Approximation (2002)
    PUB_OPTION_VAL = APPROXIMATION_AMERICAN_OPTION_FUNC(AS_VAL, PUB_STRIKE_VAL, PUB_EXPIRATION_VAL, PUB_RATE_VAL, PUB_YIELD_VAL, PUB_VOLATILITY_VAL, 1, 1, 0, 0)
Case Else 'Barone-Adesi and Whaley (1987) American approximation
    PUB_OPTION_VAL = APPROXIMATION_AMERICAN_OPTION_FUNC(AS_VAL, PUB_STRIKE_VAL, PUB_EXPIRATION_VAL, PUB_RATE_VAL, PUB_YIELD_VAL, PUB_VOLATILITY_VAL, 1, 2, 0, 0)
End Select

MGMT_OPTIONS_OBJ_FUNC = Abs(PUB_OPTION_VAL - X_VAL) ^ 2

Exit Function
ERROR_LABEL:
PUB_OPTION_VAL = 0
MGMT_OPTIONS_OBJ_FUNC = PUB_EPSILON
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : EXECUTIVE_STOCK_OPTIONS_FUNC
'DESCRIPTION   : Executive stock options
'LIBRARY       : DERIVATIVES
'GROUP         : WARRANTS
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 05/10/2011
'************************************************************************************
'************************************************************************************

Function EXECUTIVE_STOCK_OPTIONS_FUNC(ByVal SPOT As Double, _
ByVal STRIKE As Double, _
ByVal EXPIRATION As Double, _
ByVal RATE As Double, _
ByVal CARRY_COST As Double, _
ByVal SIGMA As Double, _
ByVal LAMBDA As Double, _
Optional ByVal OPTION_FLAG As Integer = 1, _
Optional ByVal CND_TYPE As Integer = 0)

Dim D1_VAL As Double
Dim D2_VAL As Double

Dim OPTION_VALUE As Double

On Error GoTo ERROR_LABEL

D1_VAL = (Log(SPOT / STRIKE) + (CARRY_COST + SIGMA ^ 2 / 2) * EXPIRATION) / (SIGMA * Sqr(EXPIRATION))
D2_VAL = D1_VAL - SIGMA * Sqr(EXPIRATION)

'------------------------------------------------------------------------------------
Select Case OPTION_FLAG
'------------------------------------------------------------------------------------
 Case 1 ', "CALL", "C"
'------------------------------------------------------------------------------------
     OPTION_VALUE = SPOT * Exp((CARRY_COST - RATE) * EXPIRATION) * CND_FUNC(D1_VAL, CND_TYPE) - STRIKE * Exp(-RATE * EXPIRATION) * CND_FUNC(D2_VAL, CND_TYPE)
'------------------------------------------------------------------------------------
 Case -1 ', "PUT", "P"
'------------------------------------------------------------------------------------
     OPTION_VALUE = STRIKE * Exp(-RATE * EXPIRATION) * CND_FUNC(-D2_VAL, CND_TYPE) - SPOT * Exp((CARRY_COST - RATE) * EXPIRATION) * CND_FUNC(-D1_VAL, CND_TYPE)
'------------------------------------------------------------------------------------
 Case Else
'------------------------------------------------------------------------------------
     GoTo ERROR_LABEL
'------------------------------------------------------------------------------------
End Select
'------------------------------------------------------------------------------------

EXECUTIVE_STOCK_OPTIONS_FUNC = Exp(-LAMBDA * EXPIRATION) * OPTION_VALUE
    
Exit Function
ERROR_LABEL:
EXECUTIVE_STOCK_OPTIONS_FUNC = Err.number
End Function
