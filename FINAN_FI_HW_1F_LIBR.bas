Attribute VB_Name = "FINAN_FI_HW_1F_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'**********************************************************************************
'**********************************************************************************
'FUNCTION      : HW_1F_CAP_FUNC
'DESCRIPTION   : HW ONE FACTOR CAP FUNCTION
'LIBRARY       : FIXED_INCOME
'GROUP         : HW_1F
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'**********************************************************************************
'**********************************************************************************

Function HW_1F_CAP_FUNC(ByVal KAPPA As Double, _
ByVal SIGMA As Double, _
ByVal FIRST_SIGMA As Double, _
ByVal SECOND_SIGMA As Double, _
ByVal RHO As Double, _
ByRef TENORS_RNG As Variant, _
ByRef ZEROS_RNG As Variant, _
ByVal STRIKE As Double, _
ByVal EXPIRATION As Double, _
ByVal CAPLET_FREQ As Double, _
ByVal NROWS As Double, _
Optional ByVal CND_TYPE As Integer = 0)
    
Dim i As Double
Dim NSIZE As Double

Dim STEMP_VAL As Double
Dim XTEMP_VAL As Double
Dim HTEMP_VAL As Double
Dim BTEMP_VAL As Double
Dim UTEMP_VAL As Double
Dim VTEMP_VAL As Double

Dim TEMP_DISC As Double
Dim TEMP_DELTA As Double
Dim TEMP_SIGMA As Double
    
Dim TEMP_VALUE As Double
Dim TEMP_RESULT As Double

Dim FIRST_DISC As Double
Dim SECOND_DISC As Double
    
Dim CAPLET_TENOR As Double
Dim FIRST_TENOR As Double
Dim SECOND_TENOR As Double
    
On Error GoTo ERROR_LABEL

'TEMP_RATE = Log(HW_DISCOUNT_FACTOR_FUNC(TENORS_RNG, ZEROS_RNG, 0, NROWS) / _
        HW_DISCOUNT_FACTOR_FUNC(TENORS_RNG, ZEROS_RNG, 0 + 0.00001, _
        NROWS)) / 0.00001

TEMP_DISC = 1
TEMP_DELTA = 1
NSIZE = (EXPIRATION * CAPLET_FREQ) - 1
CAPLET_TENOR = 1 / CAPLET_FREQ

FIRST_TENOR = EXPIRATION - CAPLET_TENOR

SECOND_TENOR = EXPIRATION
    
FIRST_DISC = HW_DISCOUNT_FACTOR_FUNC(TENORS_RNG, ZEROS_RNG, FIRST_TENOR, NROWS)
SECOND_DISC = HW_DISCOUNT_FACTOR_FUNC(TENORS_RNG, ZEROS_RNG, SECOND_TENOR, NROWS)

TEMP_RESULT = 0

For i = 1 To NSIZE
    STEMP_VAL = (1 + STRIKE * CAPLET_TENOR) * SECOND_DISC
    XTEMP_VAL = FIRST_DISC

    BTEMP_VAL = (1 - Exp(-KAPPA * (SECOND_TENOR - FIRST_TENOR))) / KAPPA
    
    UTEMP_VAL = (Exp(-KAPPA * SECOND_TENOR) - Exp(-KAPPA * FIRST_TENOR)) _
            / (KAPPA * (KAPPA - SIGMA))
    VTEMP_VAL = (Exp(-SIGMA * SECOND_TENOR) - Exp(-SIGMA * FIRST_TENOR)) _
            / (SIGMA * (KAPPA - SIGMA))

    TEMP_VALUE = FIRST_SIGMA ^ 2 / (2 * KAPPA) * BTEMP_VAL ^ 2 * _
            (1 - Exp(-2 * KAPPA * FIRST_TENOR))
    
    TEMP_VALUE = TEMP_VALUE + SECOND_SIGMA ^ 2 * (UTEMP_VAL ^ 2 / _
            (2 * KAPPA) * (Exp(2 * KAPPA * FIRST_TENOR) - 1))
    
    TEMP_VALUE = TEMP_VALUE + SECOND_SIGMA ^ 2 * (VTEMP_VAL ^ 2 / (2 * SIGMA) * _
            (Exp(2 * SIGMA * FIRST_TENOR) - 1))
    
    TEMP_VALUE = TEMP_VALUE - 2 * SECOND_SIGMA ^ 2 * UTEMP_VAL * VTEMP_VAL / _
            (KAPPA + SIGMA) * (Exp((KAPPA + SIGMA) * FIRST_TENOR) - 1)
    
    TEMP_VALUE = TEMP_VALUE + 2 * RHO * FIRST_SIGMA * SECOND_SIGMA / KAPPA * _
            (Exp(-KAPPA * FIRST_TENOR) - Exp(-KAPPA * SECOND_TENOR)) * _
            (UTEMP_VAL / (2 * KAPPA) * (Exp(2 * KAPPA * FIRST_TENOR) - 1) - _
            VTEMP_VAL / (KAPPA + SIGMA) * (Exp((KAPPA + SIGMA) * FIRST_TENOR) - 1))
    
    TEMP_SIGMA = (TEMP_VALUE) ^ 0.5
    
    HTEMP_VAL = 1 / TEMP_SIGMA * Log(STEMP_VAL / XTEMP_VAL) + 0.5 * TEMP_SIGMA
    
    TEMP_RESULT = TEMP_RESULT + XTEMP_VAL * CND_FUNC(-HTEMP_VAL + TEMP_SIGMA, CND_TYPE) - STEMP_VAL * _
            CND_FUNC(-HTEMP_VAL, CND_TYPE)
    
    SECOND_TENOR = FIRST_TENOR
    FIRST_TENOR = FIRST_TENOR - CAPLET_TENOR
    SECOND_DISC = FIRST_DISC
    FIRST_DISC = HW_DISCOUNT_FACTOR_FUNC(TENORS_RNG, ZEROS_RNG, FIRST_TENOR, NROWS)
Next i

HW_1F_CAP_FUNC = TEMP_RESULT
  
Exit Function
ERROR_LABEL:
HW_1F_CAP_FUNC = Err.number
End Function

'**********************************************************************************
'**********************************************************************************
'FUNCTION      : HW_2F_CAP_FUNC
'DESCRIPTION   : HW TWO FACTOR CAP FUNCTION
'LIBRARY       : FIXED_INCOME
'GROUP         : SWAPS
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'**********************************************************************************
'**********************************************************************************

Function HW_2F_CAP_FUNC(ByVal KAPPA As Double, _
ByVal SIGMA As Double, _
ByVal FIRST_SIGMA As Double, _
ByVal SECOND_SIGMA As Double, _
ByVal RHO As Double, _
ByVal SWAP_FREQ As Double, _
ByRef SWAP_TENOR_RNG As Variant, _
ByRef SWAP_RATE_RNG As Variant, _
ByVal CAP_FREQ As Double, _
ByRef CAP_TENOR_RNG As Variant, _
ByRef CAP_SIGMA_RNG As Variant, _
Optional ByVal INIT_RATE As Double = 0.05, _
Optional ByVal CND_TYPE As Integer = 0, _
Optional ByVal OUTPUT As Integer = 0)

'KAPPA: 10.14%
'SIGMA: 20.07%
'FIRST_SIGMA: 0.77%
'SECOND_SIGMA: 0.26%
'RHO: -29.98%
'SWAP_FREQ: 2
'CAP_FREQ: 4

Dim i As Double
Dim j As Double

Dim ii As Double
Dim jj As Double
Dim kk As Double

Dim nSWAPS As Double
Dim nCAPS As Double
Dim TEMP_SUM As Double

Dim SWAP_RATE_VECTOR As Variant
Dim SWAP_TENOR_VECTOR As Variant
Dim CAP_SIGMA_VECTOR As Variant
Dim CAP_TENOR_VECTOR As Variant
    
On Error GoTo ERROR_LABEL

SWAP_RATE_VECTOR = SWAP_RATE_RNG
If UBound(SWAP_RATE_VECTOR, 1) = 1 Then
    SWAP_RATE_VECTOR = MATRIX_TRANSPOSE_FUNC(SWAP_RATE_VECTOR)
End If

SWAP_TENOR_VECTOR = SWAP_TENOR_RNG
If UBound(SWAP_TENOR_VECTOR, 1) = 1 Then
    SWAP_TENOR_VECTOR = MATRIX_TRANSPOSE_FUNC(SWAP_TENOR_VECTOR)
End If

CAP_SIGMA_VECTOR = CAP_SIGMA_RNG
If UBound(CAP_SIGMA_VECTOR, 1) = 1 Then
    CAP_SIGMA_VECTOR = MATRIX_TRANSPOSE_FUNC(CAP_SIGMA_VECTOR)
End If

CAP_TENOR_VECTOR = CAP_TENOR_RNG
If UBound(CAP_TENOR_VECTOR, 1) = 1 Then
    CAP_TENOR_VECTOR = MATRIX_TRANSPOSE_FUNC(CAP_TENOR_VECTOR)
End If

nSWAPS = UBound(SWAP_RATE_VECTOR, 1)
nCAPS = UBound(CAP_SIGMA_VECTOR, 1)

If nSWAPS > 0 Then
      kk = nSWAPS
Else: kk = UBound(SWAP_TENOR_VECTOR, 1)
End If

ReDim HW_SWAP_TENORS(1 To kk, 1 To 1) As Variant
ReDim HW_SWAP_RATES(1 To kk, 1 To 1) As Variant

ReDim HW_ZERO_RATES(1 To kk, 1 To 1) As Variant
ReDim HW_DISC_FACTORS(1 To kk, 1 To 1) As Variant

For i = 1 To kk
    HW_SWAP_TENORS(i, 1) = SWAP_TENOR_VECTOR(i, 1)
    HW_SWAP_RATES(i, 1) = SWAP_RATE_VECTOR(i, 1)
    HW_ZERO_RATES(i, 1) = INIT_RATE
Next i

ii = HW_SWAP_TENORS(kk, 1)
ii = (ii) + 1
jj = 0

Do While HW_SWAP_TENORS(jj + 1, 1) <= 1.1
    jj = jj + 1
Loop
For i = 1 To jj
    HW_ZERO_RATES(i, 1) = 1 / HW_SWAP_TENORS(i, 1) * _
                Log(1 + HW_SWAP_TENORS(i, 1) * HW_SWAP_RATES(i, 1))
Next i

ReDim TENOR_ARR(1 To ii * SWAP_FREQ, 1 To 1) As Variant
ReDim ZERO_ARR(1 To ii * SWAP_FREQ, 1 To 1) As Variant

For i = 1 To ii * SWAP_FREQ
    TENOR_ARR(i, 1) = i / SWAP_FREQ
Next i

For i = 1 To ii * SWAP_FREQ
    TEMP_SUM = 0
    For j = 1 To (i - 1)
        TEMP_SUM = TEMP_SUM + _
        HW_ZERO_INTERPOLATION_FUNC(HW_SWAP_TENORS, HW_SWAP_RATES, _
                        (i / SWAP_FREQ), kk) _
        / SWAP_FREQ * _
                Exp(-1 * _
                HW_ZERO_INTERPOLATION_FUNC(TENOR_ARR, ZERO_ARR, _
                    (j / SWAP_FREQ), ii * SWAP_FREQ) _
                * ((j / SWAP_FREQ)))
    Next j
    If TENOR_ARR(i, 1) <= 1.1 Then
        ZERO_ARR(i, 1) = _
        HW_ZERO_INTERPOLATION_FUNC(HW_SWAP_TENORS, HW_ZERO_RATES, TENOR_ARR(i, 1), _
                        kk)
    Else
        ZERO_ARR(i, 1) = (SWAP_FREQ / (i)) * _
                    Log((1 + _
                    HW_ZERO_INTERPOLATION_FUNC(HW_SWAP_TENORS, HW_SWAP_RATES, _
                    (i / SWAP_FREQ), kk) _
                    / SWAP_FREQ) / (1 - TEMP_SUM))
    End If
Next i

For i = 1 To kk
    If HW_SWAP_TENORS(i, 1) > 1.1 Then
        HW_ZERO_RATES(i, 1) = HW_ZERO_INTERPOLATION_FUNC(TENOR_ARR, ZERO_ARR, _
                        (HW_SWAP_TENORS(i, 1)), ii * SWAP_FREQ)
    End If
    HW_DISC_FACTORS(i, 1) = Exp(-1 * HW_ZERO_RATES(i, 1) * _
                        HW_SWAP_TENORS(i, 1))
Next i


ReDim TEMP_ARR(0 To nCAPS, 1 To 3) As Variant

TEMP_ARR(0, 1) = "CAP_STRIKES"
TEMP_ARR(0, 2) = "BLACK_CAPS"
TEMP_ARR(0, 3) = "HW_CAPS"

For i = 1 To nCAPS
     
    TEMP_ARR(i, 1) = HW_FORWARD_RATES_FUNC(HW_SWAP_TENORS, HW_ZERO_RATES, _
    1 / CAP_FREQ, CAP_TENOR_VECTOR(i, 1), CAP_FREQ, kk) 'CapStrikes
'-----------------------------------------------------------------------------------------
    TEMP_ARR(i, 2) = HW_BLACK_CAP_FUNC(HW_SWAP_TENORS, HW_ZERO_RATES, _
    TEMP_ARR(i, 1), CAP_SIGMA_VECTOR(i, 1), CAP_TENOR_VECTOR(i, 1), _
    CAP_FREQ, kk, CND_TYPE) 'BlackCaps
'-----------------------------------------------------------------------------------------
    TEMP_ARR(i, 3) = HW_1F_CAP_FUNC(KAPPA, SIGMA, FIRST_SIGMA, SECOND_SIGMA, RHO, _
                    HW_SWAP_TENORS, HW_ZERO_RATES, TEMP_ARR(i, 1), _
                    CAP_TENOR_VECTOR(i, 1), CAP_FREQ, kk, CND_TYPE)
'-----------------------------------------------------------------------------------------
Next i

Select Case OUTPUT
    Case 0 'HW Table
        HW_2F_CAP_FUNC = TEMP_ARR
    Case Else 'Error
        TEMP_SUM = 0
        For i = 1 To nCAPS
            TEMP_SUM = TEMP_SUM + ((TEMP_ARR(i, 2) - TEMP_ARR(i, 3)) ^ 2)
        Next i
        HW_2F_CAP_FUNC = TEMP_SUM
End Select
    
Exit Function
ERROR_LABEL:
HW_2F_CAP_FUNC = Err.number
End Function

'**********************************************************************************
'**********************************************************************************
'FUNCTION      : HW_BLACK_CAP_FUNC

'DESCRIPTION   : HW BLACK CAP FUNCTION

'The usual and most often used model in pricing discount bond options is
'Black's 76 model which represents an extension of the famous Black-Scholes
'model. This model makes three assumptions:

'1.  that the price of the underlying variable  has a lognormal probability
'distribution at the expiration of the option

'2.  interest rates are non-stochastic and

'3.  the standard deviation of the natural logarithm of the underlying's price
'is the standard deviation of the futures/forward price of the underlying times
'the square root of the time to the option maturity

'Given these assumptions it is possible to find closed form solutions for valuing
'discount bond options and coupon bond options .

'However there are some pitfalls in using the Black 76 model. For example when
'it is used to price a cap the underlying forward rates are assumed to be
'lognormal and when it is used to price a swaption the swap rate is assumed to
'be lognormal. This shows that Black 76 inherits theoretical inconsistencies
'because both the forward rate and the swap rate cannot be distributed lognormal
'simultaneously. On the other hand Black's 76 model fails in providing solutions
'to all kinds of American options and options with more exotic payout functions.

'LIBRARY       : FIXED_INCOME
'GROUP         : HW_1F
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'**********************************************************************************
'**********************************************************************************

Function HW_BLACK_CAP_FUNC(ByRef TENORS_RNG As Variant, _
ByRef ZEROS_RNG As Variant, _
ByVal STRIKE As Double, _
ByVal SIGMA As Double, _
ByVal EXPIRATION As Double, _
ByVal CAPLET_FREQ As Double, _
ByVal NROWS As Double, _
Optional ByVal CND_TYPE As Integer = 0)

Dim i As Double
Dim NSIZE As Double

Dim D1_VAL As Double
Dim D2_VAL As Double

Dim CAPLET_TENOR As Double
Dim FIRST_TENOR As Double
Dim SECOND_TENOR As Double

Dim TEMP_DISC As Double
Dim TEMP_ROOT As Double
Dim TEMP_RATE As Double
Dim TEMP_PRICE As Double
Dim TEMP_FORWARD As Double

On Error GoTo ERROR_LABEL

TEMP_RATE = 0

NSIZE = (EXPIRATION * CAPLET_FREQ) - 1
CAPLET_TENOR = 1 / CAPLET_FREQ

FIRST_TENOR = EXPIRATION - CAPLET_TENOR
SECOND_TENOR = EXPIRATION

For i = 1 To NSIZE
    TEMP_FORWARD = HW_FORWARD_RATES_FUNC(TENORS_RNG, ZEROS_RNG, FIRST_TENOR, _
                   CAPLET_TENOR, CAPLET_FREQ, NROWS)
    
    TEMP_ROOT = SIGMA * (FIRST_TENOR) ^ 0.5
    
    D1_VAL = (Log(TEMP_FORWARD / STRIKE) + (TEMP_RATE + 0.5 * SIGMA ^ 2) * FIRST_TENOR) / TEMP_ROOT
    D2_VAL = D1_VAL - TEMP_ROOT
    
    TEMP_PRICE = TEMP_FORWARD * CND_FUNC(D1_VAL, CND_TYPE) - STRIKE * _
                 Exp(-TEMP_RATE * FIRST_TENOR) * CND_FUNC(D2_VAL, CND_TYPE)
    TEMP_DISC = HW_DISCOUNT_FACTOR_FUNC(TENORS_RNG, ZEROS_RNG, SECOND_TENOR, NROWS)
    
    HW_BLACK_CAP_FUNC = HW_BLACK_CAP_FUNC + TEMP_PRICE * TEMP_DISC * CAPLET_TENOR
    
    SECOND_TENOR = FIRST_TENOR
    FIRST_TENOR = FIRST_TENOR - CAPLET_TENOR
Next i

Exit Function
ERROR_LABEL:
HW_BLACK_CAP_FUNC = Err.number
End Function

'**********************************************************************************
'**********************************************************************************
'FUNCTION      : HW_FLOATER_FUNC
'DESCRIPTION   : With the Hull-White model it is also possible to calculate
'non-standard floating rate bonds where the interval between interest payments
'is not equal to the term of the floating rate.
'LIBRARY       : FIXED_INCOME
'GROUP         : HW_1F
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'REFERENCE: http://www.angelfire.com/ny/financeinfo/HullWhite.pdf
'**********************************************************************************
'**********************************************************************************

Function HW_FLOATER_FUNC(ByVal DELTA_ZERO As Double, _
ByVal SIGMA As Double, _
ByVal KAPPA As Double, _
ByVal FIRST_TENOR_RNG As Variant, _
ByVal DELTA_TENOR As Double, _
ByVal CASH_FLOW As Double, _
ByRef TENOR_RNG As Variant, _
ByRef ZERO_RNG As Variant)

Dim i As Double
Dim j As Double
Dim k As Double
Dim l As Double

Dim ii As Double
Dim jj As Double

Dim JMAX_VAL As Double
Dim JMIN_VAL As Double

Dim JPLUS_VAL As Double
Dim JMINUS_VAL As Double

Dim NSIZE As Double
Dim NROWS As Double
Dim nSTEPS As Double

Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double
Dim CTEMP_VAL As Double
Dim DTEMP_VAL As Double

Dim TEMP_SUM As Double
Dim DELTA_CASH As Double

Dim Q_ARR As Variant
Dim O_ARR As Variant
Dim R_ARR As Variant

Dim ALPHA_VECTOR As Variant
Dim STEPS_VECTOR As Variant
Dim FIRST_TENOR_VECTOR As Variant

On Error GoTo ERROR_LABEL

If IsArray(FIRST_TENOR_RNG) = True Then
    FIRST_TENOR_VECTOR = FIRST_TENOR_RNG
    If UBound(FIRST_TENOR_VECTOR, 1) = 1 Then: _
        FIRST_TENOR_VECTOR = MATRIX_TRANSPOSE_FUNC(FIRST_TENOR_VECTOR)
Else
    ReDim FIRST_TENOR_VECTOR(1 To 1, 1 To 1)
    FIRST_TENOR_VECTOR(1, 1) = FIRST_TENOR_RNG
End If

NROWS = UBound(FIRST_TENOR_VECTOR, 1)

'dimensioning variables
If FIRST_TENOR_VECTOR(NROWS, 1) / DELTA_ZERO - _
        Int(FIRST_TENOR_VECTOR(NROWS, 1) / DELTA_ZERO) = 0 Then
            nSTEPS = FIRST_TENOR_VECTOR(NROWS, 1) / DELTA_ZERO
Else:       nSTEPS = Int(FIRST_TENOR_VECTOR(NROWS, 1) / DELTA_ZERO) + 1
End If


'calculating constants
ATEMP_VAL = SIGMA ^ 2 * (1 - Exp(-2 * KAPPA * DELTA_ZERO)) / (2 * KAPPA)
BTEMP_VAL = Sqr(3 * ATEMP_VAL)
JMAX_VAL = Int(0.184 / (KAPPA * DELTA_ZERO)) + 1
JMIN_VAL = -JMAX_VAL
NSIZE = MINIMUM_FUNC(nSTEPS, JMAX_VAL)

'redimensioning variables
ReDim ALPHA_VECTOR(0 To nSTEPS)
ReDim Q_ARR(0 To nSTEPS + 1, JMIN_VAL To JMAX_VAL)
ReDim O_ARR(0 To nSTEPS, JMIN_VAL To JMAX_VAL)
ReDim R_ARR(0 To nSTEPS, JMIN_VAL To JMAX_VAL)
ReDim STEPS_VECTOR(0 To NROWS)

'calculating start values
ALPHA_VECTOR(0) = -Log(HW_PRICE_FUNC(TENOR_RNG, ZERO_RNG, DELTA_ZERO)) / DELTA_ZERO
R_ARR(0, 0) = ALPHA_VECTOR(0)
Q_ARR(1, 1) = HW_PROBABILITIES_TREE_FUNC(1, 0, JMAX_VAL, KAPPA, DELTA_ZERO) * _
    Exp(-R_ARR(0, 0) * DELTA_ZERO)
Q_ARR(1, 0) = HW_PROBABILITIES_TREE_FUNC(0, 0, JMAX_VAL, KAPPA, DELTA_ZERO) * _
    Exp(-R_ARR(0, 0) * DELTA_ZERO)
Q_ARR(1, -1) = HW_PROBABILITIES_TREE_FUNC(-1, 0, JMAX_VAL, KAPPA, DELTA_ZERO) * _
    Exp(-R_ARR(0, 0) * DELTA_ZERO)

JMINUS_VAL = 0
JPLUS_VAL = 0

ii = -1
jj = 1

For l = 1 To NSIZE - 1
    JMINUS_VAL = JMINUS_VAL - 1
    JPLUS_VAL = JPLUS_VAL + 1
    TEMP_SUM = 0
    For j = JMINUS_VAL To JPLUS_VAL
        TEMP_SUM = TEMP_SUM + Q_ARR(l, j) * Exp(-j * BTEMP_VAL * DELTA_ZERO)
    Next j
    ALPHA_VECTOR(l) = (Log(TEMP_SUM) - Log(HW_PRICE_FUNC(TENOR_RNG, _
            ZERO_RNG, (l + 1) * DELTA_ZERO))) / DELTA_ZERO
    For j = JMINUS_VAL To JPLUS_VAL
        R_ARR(l, j) = ALPHA_VECTOR(l) + j * BTEMP_VAL
    Next j
    ii = ii - 1
    jj = jj + 1
    For j = ii To jj
        If j = jj Then
            Q_ARR(l + 1, j) = Q_ARR(l, j - 1) * HW_PROBABILITIES_TREE_FUNC(1, j - 1, _
                        JMAX_VAL, KAPPA, DELTA_ZERO) * Exp(-R_ARR(l, j - 1) * DELTA_ZERO)
        ElseIf j = jj - 1 Then
            Q_ARR(l + 1, j) = Q_ARR(l, j) * HW_PROBABILITIES_TREE_FUNC(0, j, JMAX_VAL, KAPPA, _
                        DELTA_ZERO) * Exp(-R_ARR(l, j) * DELTA_ZERO) + Q_ARR(l, j - 1) _
                        * HW_PROBABILITIES_TREE_FUNC(1, j - 1, JMAX_VAL, KAPPA, DELTA_ZERO) * _
                        Exp(-R_ARR(l, j - 1) * DELTA_ZERO)
        ElseIf j < jj - 1 And j > ii + 1 Then
            Q_ARR(l + 1, j) = Q_ARR(l, j - 1) * HW_PROBABILITIES_TREE_FUNC(1, j - 1, JMAX_VAL, _
                        KAPPA, DELTA_ZERO) * Exp(-R_ARR(l, j - 1) * DELTA_ZERO) + _
                        Q_ARR(l, j) * HW_PROBABILITIES_TREE_FUNC(0, j, JMAX_VAL, KAPPA, DELTA_ZERO) * _
                        Exp(-R_ARR(l, j) * DELTA_ZERO) + Q_ARR(l, j + 1) * _
                        HW_PROBABILITIES_TREE_FUNC(-1, j + 1, JMAX_VAL, KAPPA, DELTA_ZERO) * _
                        Exp(-R_ARR(l, j + 1) * DELTA_ZERO)
        ElseIf j = ii + 1 Then
            Q_ARR(l + 1, j) = Q_ARR(l, j) * HW_PROBABILITIES_TREE_FUNC(0, j, JMAX_VAL, KAPPA, _
                        DELTA_ZERO) * Exp(-R_ARR(l, j) * DELTA_ZERO) + Q_ARR(l, _
                        j + 1) * HW_PROBABILITIES_TREE_FUNC(-1, j + 1, JMAX_VAL, KAPPA, DELTA_ZERO) * _
                        Exp(-R_ARR(l, j + 1) * DELTA_ZERO)
        ElseIf j = ii Then
            Q_ARR(l + 1, j) = Q_ARR(l, j + 1) * HW_PROBABILITIES_TREE_FUNC(-1, j + 1, JMAX_VAL, _
                        KAPPA, DELTA_ZERO) * Exp(-R_ARR(l, j + 1) * DELTA_ZERO)
        End If
    Next j
Next l
    
    'using the Q_ARR(NSIZE) to calculate ALPHA_VECTOR(NSIZE) and so on up to NSTEPS
For l = NSIZE To nSTEPS
    
    JPLUS_VAL = JMAX_VAL
    JMINUS_VAL = JMIN_VAL
    
    TEMP_SUM = 0
    For j = JMIN_VAL To JMAX_VAL
        TEMP_SUM = TEMP_SUM + Q_ARR(l, j) * Exp(-j * BTEMP_VAL * DELTA_ZERO)
    Next j
    
    ALPHA_VECTOR(l) = (Log(TEMP_SUM) - Log(HW_PRICE_FUNC(TENOR_RNG, _
                    ZERO_RNG, (l + 1) * DELTA_ZERO))) / DELTA_ZERO
    For j = JMINUS_VAL To JPLUS_VAL
        R_ARR(l, j) = ALPHA_VECTOR(l) + j * BTEMP_VAL
    Next j
    
    For j = JMINUS_VAL To JPLUS_VAL
    'considering the spezial case when nonstandard branching leads to central
    'nodes with five incoming arrows
        If JMAX_VAL = 2 Then
            Q_ARR(l + 1, 2) = Q_ARR(l, 2) * HW_PROBABILITIES_TREE_FUNC(1, 2, JMAX_VAL, KAPPA, _
                            DELTA_ZERO) * Exp(-R_ARR(l, 2) * DELTA_ZERO) + _
                            Q_ARR(l, 1) * HW_PROBABILITIES_TREE_FUNC(1, 1, JMAX_VAL, KAPPA, _
                            DELTA_ZERO) * Exp(-R_ARR(l, 1) * DELTA_ZERO)
            Q_ARR(l + 1, 1) = Q_ARR(l, 2) * HW_PROBABILITIES_TREE_FUNC(0, 2, JMAX_VAL, KAPPA, _
                            DELTA_ZERO) * Exp(-R_ARR(l, 2) * DELTA_ZERO) + _
                            Q_ARR(l, 1) * HW_PROBABILITIES_TREE_FUNC(0, 1, JMAX_VAL, KAPPA, _
                            DELTA_ZERO) * Exp(-R_ARR(l, 1) * DELTA_ZERO) + _
                            Q_ARR(l, 0) * HW_PROBABILITIES_TREE_FUNC(1, 0, JMAX_VAL, KAPPA, _
                            DELTA_ZERO) * Exp(-R_ARR(l, 0) * DELTA_ZERO)
            Q_ARR(l + 1, 0) = Q_ARR(l, 2) * HW_PROBABILITIES_TREE_FUNC(-1, 2, JMAX_VAL, KAPPA, _
                            DELTA_ZERO) * Exp(-R_ARR(l, 2) * DELTA_ZERO) + _
                            Q_ARR(l, 1) * HW_PROBABILITIES_TREE_FUNC(-1, 1, JMAX_VAL, KAPPA, _
                            DELTA_ZERO) * Exp(-R_ARR(l, 1) * DELTA_ZERO) + _
                            Q_ARR(l, 0) * HW_PROBABILITIES_TREE_FUNC(0, 0, JMAX_VAL, KAPPA, _
                            DELTA_ZERO) * Exp(-R_ARR(l, 0) * DELTA_ZERO) + _
                            Q_ARR(l, -1) * HW_PROBABILITIES_TREE_FUNC(1, -1, JMAX_VAL, KAPPA, _
                            DELTA_ZERO) * Exp(-R_ARR(l, -1) * DELTA_ZERO) + _
                            Q_ARR(l, -2) * HW_PROBABILITIES_TREE_FUNC(1, -2, JMAX_VAL, KAPPA, _
                            DELTA_ZERO) * Exp(-R_ARR(l, -2) * DELTA_ZERO)
            Q_ARR(l + 1, -1) = Q_ARR(l, 0) * HW_PROBABILITIES_TREE_FUNC(-1, 0, JMAX_VAL, KAPPA, _
                            DELTA_ZERO) * Exp(-R_ARR(l, 0) * DELTA_ZERO) + _
                            Q_ARR(l, -1) * HW_PROBABILITIES_TREE_FUNC(0, -1, JMAX_VAL, KAPPA, _
                            DELTA_ZERO) * Exp(-R_ARR(l, -1) * DELTA_ZERO) + _
                            Q_ARR(l, -2) * HW_PROBABILITIES_TREE_FUNC(1, -2, JMAX_VAL, KAPPA, _
                            DELTA_ZERO) * Exp(-R_ARR(l, -2) * DELTA_ZERO)
            Q_ARR(l + 1, -2) = Q_ARR(l, -1) * HW_PROBABILITIES_TREE_FUNC(-1, -1, JMAX_VAL, KAPPA, _
                            DELTA_ZERO) * Exp(-R_ARR(l, -1) * DELTA_ZERO) + _
                            Q_ARR(l, -2) * HW_PROBABILITIES_TREE_FUNC(-1, -2, JMAX_VAL, KAPPA, _
                            DELTA_ZERO) * Exp(-R_ARR(l, -2) * DELTA_ZERO)
        Else
            If j = JPLUS_VAL Then
                Q_ARR(l + 1, j) = Q_ARR(l, j) * HW_PROBABILITIES_TREE_FUNC(1, j, JMAX_VAL, KAPPA, _
                            DELTA_ZERO) * Exp(-R_ARR(l, j) * DELTA_ZERO) + _
                            Q_ARR(l, j - 1) * HW_PROBABILITIES_TREE_FUNC(1, j - 1, JMAX_VAL, _
                            KAPPA, DELTA_ZERO) * Exp(-R_ARR(l, j - 1) * DELTA_ZERO)
            ElseIf j = JPLUS_VAL - 1 Then
                Q_ARR(l + 1, j) = Q_ARR(l, j + 1) * HW_PROBABILITIES_TREE_FUNC(0, j + 1, JMAX_VAL, _
                            KAPPA, DELTA_ZERO) * Exp(-R_ARR(l, j + 1) * DELTA_ZERO) _
                            + Q_ARR(l, j) * HW_PROBABILITIES_TREE_FUNC(0, j, JMAX_VAL, KAPPA, _
                            DELTA_ZERO) * Exp(-R_ARR(l, j) * DELTA_ZERO) + _
                            Q_ARR(l, j - 1) * HW_PROBABILITIES_TREE_FUNC(1, j - 1, JMAX_VAL, _
                            KAPPA, DELTA_ZERO) * Exp(-R_ARR(l, j - 1) * DELTA_ZERO)
            ElseIf j = JPLUS_VAL - 2 Then
                Q_ARR(l + 1, j) = Q_ARR(l, j + 1) * HW_PROBABILITIES_TREE_FUNC(-1, j + 1, JMAX_VAL, _
                            KAPPA, DELTA_ZERO) * Exp(-R_ARR(l, j + 1) * DELTA_ZERO) _
                            + Q_ARR(l, j) * HW_PROBABILITIES_TREE_FUNC(0, j, JMAX_VAL, KAPPA, _
                            DELTA_ZERO) * Exp(-R_ARR(l, j) * DELTA_ZERO) + _
                            Q_ARR(l, j - 1) * HW_PROBABILITIES_TREE_FUNC(1, j - 1, JMAX_VAL, _
                            KAPPA, DELTA_ZERO) * Exp(-R_ARR(l, j - 1) * _
                            DELTA_ZERO) + Q_ARR(l, j + 2) * HW_PROBABILITIES_TREE_FUNC(-1, _
                            j + 2, JMAX_VAL, KAPPA, DELTA_ZERO) * Exp(-R_ARR(l, j + 2) _
                            * DELTA_ZERO)
            ElseIf j < JPLUS_VAL - 2 And j > JMINUS_VAL + 2 Then
                Q_ARR(l + 1, j) = Q_ARR(l, j + 1) * HW_PROBABILITIES_TREE_FUNC(-1, j + 1, JMAX_VAL, _
                            KAPPA, DELTA_ZERO) * Exp(-R_ARR(l, j + 1) * DELTA_ZERO) _
                            + Q_ARR(l, j) * HW_PROBABILITIES_TREE_FUNC(0, j, JMAX_VAL, KAPPA, _
                            DELTA_ZERO) * Exp(-R_ARR(l, j) * DELTA_ZERO) + _
                            Q_ARR(l, j - 1) * HW_PROBABILITIES_TREE_FUNC(1, j - 1, JMAX_VAL, _
                            KAPPA, DELTA_ZERO) * Exp(-R_ARR(l, j - 1) * DELTA_ZERO)
            ElseIf j = JMINUS_VAL + 2 Then
                Q_ARR(l + 1, j) = Q_ARR(l, j + 1) * HW_PROBABILITIES_TREE_FUNC(-1, j + 1, JMAX_VAL, _
                            KAPPA, DELTA_ZERO) * Exp(-R_ARR(l, j + 1) * DELTA_ZERO) _
                            + Q_ARR(l, j) * HW_PROBABILITIES_TREE_FUNC(0, j, JMAX_VAL, KAPPA, _
                            DELTA_ZERO) * Exp(-R_ARR(l, j) * DELTA_ZERO) + _
                            Q_ARR(l, j - 1) * HW_PROBABILITIES_TREE_FUNC(1, j - 1, JMAX_VAL, KAPPA, _
                            DELTA_ZERO) * Exp(-R_ARR(l, j - 1) * DELTA_ZERO) + _
                            Q_ARR(l, j - 2) * HW_PROBABILITIES_TREE_FUNC(1, j - 2, JMAX_VAL, KAPPA, _
                            DELTA_ZERO) * Exp(-R_ARR(l, j - 2) * DELTA_ZERO)
            ElseIf j = JMINUS_VAL + 1 Then
                Q_ARR(l + 1, j) = Q_ARR(l, j) * HW_PROBABILITIES_TREE_FUNC(0, j, JMAX_VAL, KAPPA, _
                            DELTA_ZERO) * Exp(-R_ARR(l, j) * DELTA_ZERO) + _
                            Q_ARR(l, j - 1) * HW_PROBABILITIES_TREE_FUNC(0, j - 1, JMAX_VAL, _
                            KAPPA, DELTA_ZERO) * Exp(-R_ARR(l, j - 1) * DELTA_ZERO) _
                            + Q_ARR(l, j + 1) * HW_PROBABILITIES_TREE_FUNC(-1, j + 1, JMAX_VAL, KAPPA, _
                            DELTA_ZERO) * Exp(-R_ARR(l, j + 1) * DELTA_ZERO)
            ElseIf j = JMINUS_VAL Then
                Q_ARR(l + 1, j) = Q_ARR(l, j) * HW_PROBABILITIES_TREE_FUNC(-1, j, JMAX_VAL, KAPPA, _
                            DELTA_ZERO) * Exp(-R_ARR(l, j) * DELTA_ZERO) + Q_ARR(l, _
                            j + 1) * HW_PROBABILITIES_TREE_FUNC(-1, j + 1, JMAX_VAL, KAPPA, DELTA_ZERO) _
                            * Exp(-R_ARR(l, j + 1) * DELTA_ZERO)
            End If
        End If
    Next j
Next l

'cash flow calculations
For k = 1 To NROWS
    If k = NROWS Then DELTA_CASH = CASH_FLOW
        If k = 1 Then
            DTEMP_VAL = FIRST_TENOR_VECTOR(k, 1)
        Else: DTEMP_VAL = FIRST_TENOR_VECTOR(k, 1) - FIRST_TENOR_VECTOR(k - 1, 1)
        End If
        CTEMP_VAL = FIRST_TENOR_VECTOR(k, 1) / DELTA_ZERO - _
                Int(FIRST_TENOR_VECTOR(k, 1) / DELTA_ZERO)
    If CTEMP_VAL = 0 Then
        STEPS_VECTOR(k) = FIRST_TENOR_VECTOR(k, 1) / DELTA_ZERO
        For j = -MINIMUM_FUNC(STEPS_VECTOR(k), JMAX_VAL) To MINIMUM_FUNC(STEPS_VECTOR(k), JMAX_VAL)
            O_ARR(STEPS_VECTOR(k), j) = (DELTA_CASH + CASH_FLOW * (Exp(DTEMP_VAL * _
                                HW_ZERO_FUNC(HW_CONTINUOUS_RATE_FUNC(R_ARR(STEPS_VECTOR(k), j), _
                                DELTA_ZERO, FIRST_TENOR_VECTOR(k, 1), SIGMA, KAPPA, _
                                TENOR_RNG, ZERO_RNG), FIRST_TENOR_VECTOR(k, 1), _
                                FIRST_TENOR_VECTOR(k, 1) + DELTA_TENOR, SIGMA, KAPPA, _
                                TENOR_RNG, ZERO_RNG)) - 1)) * Exp(-DTEMP_VAL * _
                                HW_ZERO_FUNC(HW_CONTINUOUS_RATE_FUNC(R_ARR(STEPS_VECTOR(k), j), _
                                DELTA_ZERO, FIRST_TENOR_VECTOR(k, 1), SIGMA, KAPPA, _
                                TENOR_RNG, ZERO_RNG), FIRST_TENOR_VECTOR(k, 1), _
                                FIRST_TENOR_VECTOR(k, 1) + DTEMP_VAL, SIGMA, KAPPA, _
                                TENOR_RNG, ZERO_RNG)) + O_ARR(STEPS_VECTOR(k), j)
        Next j
    Else
        STEPS_VECTOR(k) = Int(FIRST_TENOR_VECTOR(k, 1) / DELTA_ZERO)
        For j = -MINIMUM_FUNC(STEPS_VECTOR(k), JMAX_VAL) To MINIMUM_FUNC(STEPS_VECTOR(k), JMAX_VAL)
            O_ARR(STEPS_VECTOR(k), j) = CTEMP_VAL * CASH_FLOW * DTEMP_VAL * _
                                HW_ZERO_FUNC(HW_CONTINUOUS_RATE_FUNC(R_ARR(STEPS_VECTOR(k), j), _
                                DELTA_ZERO, STEPS_VECTOR(k) * DELTA_ZERO, SIGMA, KAPPA, _
                                TENOR_RNG, ZERO_RNG), STEPS_VECTOR(k) * DELTA_ZERO, _
                                STEPS_VECTOR(k) * DELTA_ZERO + DELTA_TENOR, SIGMA, _
                                KAPPA, TENOR_RNG, ZERO_RNG) + O_ARR(STEPS_VECTOR(k), j)
        Next j
        For j = -MINIMUM_FUNC(STEPS_VECTOR(k) + 1, JMAX_VAL) To _
            MINIMUM_FUNC(STEPS_VECTOR(k) + 1, JMAX_VAL)
                O_ARR(STEPS_VECTOR(k) + 1, j) = (1 - CTEMP_VAL) * CASH_FLOW * DTEMP_VAL * _
                            HW_ZERO_FUNC(HW_CONTINUOUS_RATE_FUNC(R_ARR(STEPS_VECTOR(k) + 1, j), _
                            DELTA_ZERO, (STEPS_VECTOR(k) + 1) * DELTA_ZERO, SIGMA, _
                            KAPPA, TENOR_RNG, ZERO_RNG), (STEPS_VECTOR(k) + 1) * _
                            DELTA_ZERO, (STEPS_VECTOR(k) + 1) * DELTA_ZERO + DELTA_TENOR, _
                            SIGMA, KAPPA, TENOR_RNG, ZERO_RNG) + O_ARR(STEPS_VECTOR(k) + _
                            1, j)
        Next j
    End If
Next k

'backward induction
For i = nSTEPS To NSIZE + 1 Step -1
    For j = JMIN_VAL To JMAX_VAL
        If j = JMAX_VAL Then
            O_ARR(i - 1, j) = (O_ARR(i, j) * HW_PROBABILITIES_TREE_FUNC(1, j, JMAX_VAL, _
                            KAPPA, DELTA_ZERO) + O_ARR(i, j - 1) * HW_PROBABILITIES_TREE_FUNC(0, j, _
                            JMAX_VAL, KAPPA, DELTA_ZERO) + O_ARR(i, j - 2) * HW_PROBABILITIES_TREE_FUNC(-1, _
                            j, JMAX_VAL, KAPPA, DELTA_ZERO)) * Exp(-R_ARR(i - 1, j) * _
                            DELTA_ZERO) + O_ARR(i - 1, j)
        ElseIf j < JMAX_VAL And j > JMIN_VAL Then
            O_ARR(i - 1, j) = (O_ARR(i, j + 1) * HW_PROBABILITIES_TREE_FUNC(1, j, JMAX_VAL, _
                            KAPPA, DELTA_ZERO) + O_ARR(i, j) * HW_PROBABILITIES_TREE_FUNC(0, j, JMAX_VAL, _
                            KAPPA, DELTA_ZERO) + O_ARR(i, j - 1) * HW_PROBABILITIES_TREE_FUNC(-1, j, _
                            JMAX_VAL, KAPPA, DELTA_ZERO)) * Exp(-R_ARR(i - 1, j) * DELTA_ZERO) _
                            + O_ARR(i - 1, j)
        ElseIf j = JMIN_VAL Then
            O_ARR(i - 1, j) = (O_ARR(i, j) * HW_PROBABILITIES_TREE_FUNC(-1, j, JMAX_VAL, KAPPA, _
                            DELTA_ZERO) + O_ARR(i, j + 1) * HW_PROBABILITIES_TREE_FUNC(0, j, JMAX_VAL, _
                            KAPPA, DELTA_ZERO) + O_ARR(i, j + 2) * HW_PROBABILITIES_TREE_FUNC(1, j, _
                            JMAX_VAL, KAPPA, DELTA_ZERO)) * Exp(-R_ARR(i - 1, j) * _
                            DELTA_ZERO) + O_ARR(i - 1, j)
        End If
    Next j
Next i

JMINUS_VAL = -NSIZE
JPLUS_VAL = NSIZE

For i = NSIZE To 1 Step -1
    JMINUS_VAL = JMINUS_VAL + 1
    JPLUS_VAL = JPLUS_VAL - 1
    For j = JMINUS_VAL To JPLUS_VAL
        O_ARR(i - 1, j) = (O_ARR(i, j + 1) * HW_PROBABILITIES_TREE_FUNC(1, j, JMAX_VAL, KAPPA, _
                        DELTA_ZERO) + O_ARR(i, j) * HW_PROBABILITIES_TREE_FUNC(0, j, JMAX_VAL, _
                        KAPPA, DELTA_ZERO) + O_ARR(i, j - 1) * HW_PROBABILITIES_TREE_FUNC(-1, j, _
                        JMAX_VAL, KAPPA, DELTA_ZERO)) * Exp(-R_ARR(i - 1, j) * _
                        DELTA_ZERO) + O_ARR(i - 1, j)
    Next j
Next i


HW_FLOATER_FUNC = O_ARR(0, 0)
   
Exit Function
ERROR_LABEL:
    HW_FLOATER_FUNC = Err.number
End Function


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : HW_CAP_FLOOR_ANALYTICAL_FUNC
'DESCRIPTION   : Valuation of Caps and Floors, comparison of analytical and
'numerical solution
'LIBRARY       : FIXED_INCOME
'GROUP         : HW_1F
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'REFERENCE: http://www.angelfire.com/ny/financeinfo/HullWhite.pdf
'**********************************************************************************
'**********************************************************************************


Function HW_CAP_FLOOR_ANALYTICAL_FUNC(ByVal OPTION_FLAG As Integer, _
ByVal SIGMA As Double, _
ByVal KAPPA As Double, _
ByVal TAU_VAL As Double, _
ByVal STRIKE_RATE As Double, _
ByVal SECOND_TENOR As Double, _
ByVal CASH_FLOW As Double, _
ByRef TENOR_RNG As Variant, _
ByRef ZERO_RNG As Variant, _
Optional ByVal CND_TYPE As Integer = 0)

'Cap valuation:
'payout function:
'max(l(1-exp(tau(RcX-Rck);0)
'Rck...cont. comp. floating rate which applies to a period k
'RcX....cont. comp. strike rate

Dim i As Double
Dim TEMP_SUM As Double
Dim DELTA_CASH As Double
Dim DELTA_TENOR As Double

On Error GoTo ERROR_LABEL

TEMP_SUM = 0
For i = TAU_VAL To SECOND_TENOR - TAU_VAL Step TAU_VAL
    DELTA_TENOR = i + TAU_VAL
    DELTA_CASH = CASH_FLOW * Exp(TAU_VAL * STRIKE_RATE)
    TEMP_SUM = TEMP_SUM + HW_EURO_BOND_VALUATION_FUNC(OPTION_FLAG, _
                DELTA_TENOR, DELTA_CASH, CASH_FLOW, KAPPA, _
                SIGMA, i, TENOR_RNG, ZERO_RNG, CND_TYPE)
Next i

HW_CAP_FLOOR_ANALYTICAL_FUNC = TEMP_SUM

Exit Function
ERROR_LABEL:
HW_CAP_FLOOR_ANALYTICAL_FUNC = Err.number
End Function


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : HW_CAP_FLOOR_FUNC
'DESCRIPTION   : A cap provides a payment everytime a specified floating rate like the 6 month
'LIBOR exceeds the agreed cap rate. This assures the borrower of a loan that
'he will never pay more than the cap rate. These individual payments are called
'caplets and as a sum make up a cap.

'Usually the reference rate is observed at the beginning of the period starting
'at t=1 and eventually a payment is made at the end of the period. In order to
'use the Hull-White model which determines the reference rate at each point in
'time and state we have to calculate the cash flows at the time the reference
'rate is measured which is at the beginning of the period.

'To evaluate a floor z takes on -1 and either the numerical procedure or the
'replicating approach of call options on discount bonds is used. A collar is
'simply a combination of a long position in a cap and a short position in a
'floor and is calculated accordingly.

'LIBRARY       : FIXED_INCOME
'GROUP         : HW_1F
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'REFERENCE: http://www.angelfire.com/ny/financeinfo/HullWhite.pdf
'**********************************************************************************
'**********************************************************************************

Function HW_CAP_FLOOR_FUNC(ByVal OPTION_FLAG As Integer, _
ByVal nSTEPS As Double, _
ByVal SIGMA As Double, _
ByVal KAPPA As Double, _
ByVal TAU_VAL As Double, _
ByVal DELTA_TENOR As Double, _
ByVal SECOND_TENOR As Double, _
ByVal CASH_FLOW As Double, _
ByVal STRIKE As Double, _
ByRef TENOR_RNG As Variant, _
ByRef ZERO_RNG As Variant)

Dim i As Double
Dim TEMP_SUM As Double

On Error GoTo ERROR_LABEL

TEMP_SUM = 0

For i = TAU_VAL To SECOND_TENOR - TAU_VAL Step TAU_VAL
    TEMP_SUM = TEMP_SUM + HW_CAP_FLOOR_LET_FUNC(OPTION_FLAG, nSTEPS, SIGMA, KAPPA, i, _
            TAU_VAL, DELTA_TENOR, SECOND_TENOR, CASH_FLOW, STRIKE, TENOR_RNG, ZERO_RNG)
Next i
HW_CAP_FLOOR_FUNC = TEMP_SUM

Exit Function
ERROR_LABEL:
    HW_CAP_FLOOR_FUNC = Err.number
End Function


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : HW_CAP_FLOOR_LET_FUNC
'DESCRIPTION   : Valuation of standard and non-standard floater
'LIBRARY       : FIXED_INCOME
'GROUP         : HW_1F
'ID            : 007
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'REFERENCE: http://www.angelfire.com/ny/financeinfo/HullWhite.pdf
'**********************************************************************************
'**********************************************************************************

Function HW_CAP_FLOOR_LET_FUNC(ByVal OPTION_FLAG As Integer, _
ByVal nSTEPS As Double, _
ByVal SIGMA As Double, _
ByVal KAPPA As Double, _
ByVal FIRST_TENOR As Double, _
ByVal TAU_VAL As Double, _
ByVal DELTA_TENOR As Double, _
ByVal SECOND_TENOR As Double, _
ByVal CASH_FLOW As Double, _
ByVal STRIKE As Double, _
ByRef TENOR_RNG As Variant, _
ByRef ZERO_RNG As Variant)

Dim h As Double
Dim i As Double
Dim j As Double
Dim l As Double

Dim ii As Double
Dim jj As Double

Dim JMAX_VAL As Double
Dim JMIN_VAL As Double

Dim JPLUS_VAL As Double
Dim JMINUS_VAL As Double

Dim NSIZE As Double

Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double

Dim TEMP_SUM As Double
Dim TEMP_DELTA As Double

Dim Q_ARR As Variant
Dim O_ARR As Variant
Dim R_ARR As Variant

Dim ALPHA_VECTOR As Variant

On Error GoTo ERROR_LABEL

Select Case OPTION_FLAG
    Case -1 '"p", "put", -1
        h = 1
    Case Else
        h = -1
End Select

TEMP_DELTA = FIRST_TENOR / nSTEPS
JMAX_VAL = Int(0.184 / (KAPPA * TEMP_DELTA)) + 1
JMIN_VAL = -JMAX_VAL
NSIZE = MINIMUM_FUNC(nSTEPS, JMAX_VAL)

ReDim ALPHA_VECTOR(0 To nSTEPS)
ReDim Q_ARR(0 To nSTEPS + 1, -NSIZE To NSIZE)
ReDim O_ARR(0 To nSTEPS, -NSIZE To NSIZE)
ReDim R_ARR(0 To nSTEPS, -NSIZE To NSIZE)

ATEMP_VAL = SIGMA ^ 2 * (1 - Exp(-2 * KAPPA * TEMP_DELTA)) / (2 * KAPPA)
BTEMP_VAL = Sqr(3 * ATEMP_VAL)
JMINUS_VAL = 1
JPLUS_VAL = -1

ALPHA_VECTOR(0) = -Log(HW_PRICE_FUNC(TENOR_RNG, ZERO_RNG, TEMP_DELTA)) / TEMP_DELTA
R_ARR(0, 0) = ALPHA_VECTOR(0)
Q_ARR(1, 1) = HW_PROBABILITIES_TREE_FUNC(1, 0, JMAX_VAL, KAPPA, TEMP_DELTA) * Exp(-R_ARR(0, 0) * TEMP_DELTA)
Q_ARR(1, 0) = HW_PROBABILITIES_TREE_FUNC(0, 0, JMAX_VAL, KAPPA, TEMP_DELTA) * Exp(-R_ARR(0, 0) * TEMP_DELTA)
Q_ARR(1, -1) = HW_PROBABILITIES_TREE_FUNC(-1, 0, JMAX_VAL, KAPPA, TEMP_DELTA) * Exp(-R_ARR(0, 0) * TEMP_DELTA)

JMINUS_VAL = 0
JPLUS_VAL = 0

ii = -1
jj = 1

For l = 1 To NSIZE - 1
    JMINUS_VAL = JMINUS_VAL - 1
    JPLUS_VAL = JPLUS_VAL + 1
    
    TEMP_SUM = 0
    For j = JMINUS_VAL To JPLUS_VAL
        TEMP_SUM = TEMP_SUM + Q_ARR(l, j) * Exp(-j * BTEMP_VAL * TEMP_DELTA)
    Next j
    
    ALPHA_VECTOR(l) = (Log(TEMP_SUM) - Log(HW_PRICE_FUNC(TENOR_RNG, _
            ZERO_RNG, (l + 1) * TEMP_DELTA))) / TEMP_DELTA
    For j = JMINUS_VAL To JPLUS_VAL
        R_ARR(l, j) = ALPHA_VECTOR(l) + j * BTEMP_VAL
    Next j
    
    ii = ii - 1
    jj = jj + 1
    
    For j = ii To jj
    
        If j = jj Then
            Q_ARR(l + 1, j) = Q_ARR(l, j - 1) * HW_PROBABILITIES_TREE_FUNC(1, j - 1, JMAX_VAL, _
                    KAPPA, TEMP_DELTA) * Exp(-R_ARR(l, j - 1) * TEMP_DELTA)
        ElseIf j = jj - 1 Then
            Q_ARR(l + 1, j) = Q_ARR(l, j) * HW_PROBABILITIES_TREE_FUNC(0, j, JMAX_VAL, KAPPA, _
                    TEMP_DELTA) * Exp(-R_ARR(l, j) * TEMP_DELTA) + Q_ARR(l, j - 1) * _
                    HW_PROBABILITIES_TREE_FUNC(1, j - 1, JMAX_VAL, KAPPA, TEMP_DELTA) * Exp(-R_ARR(l, _
                    j - 1) * TEMP_DELTA)
        ElseIf j < jj - 1 And j > ii + 1 Then
            Q_ARR(l + 1, j) = Q_ARR(l, j - 1) * HW_PROBABILITIES_TREE_FUNC(1, j - 1, JMAX_VAL, _
                    KAPPA, TEMP_DELTA) * Exp(-R_ARR(l, j - 1) * TEMP_DELTA) + _
                    Q_ARR(l, j) * HW_PROBABILITIES_TREE_FUNC(0, j, JMAX_VAL, KAPPA, TEMP_DELTA) * _
                    Exp(-R_ARR(l, j) * TEMP_DELTA) + Q_ARR(l, j + 1) * HW_PROBABILITIES_TREE_FUNC(-1, _
                    j + 1, JMAX_VAL, KAPPA, TEMP_DELTA) * Exp(-R_ARR(l, j + 1) * TEMP_DELTA)
        ElseIf j = ii + 1 Then
            Q_ARR(l + 1, j) = Q_ARR(l, j) * HW_PROBABILITIES_TREE_FUNC(0, j, JMAX_VAL, KAPPA, _
                    TEMP_DELTA) * Exp(-R_ARR(l, j) * TEMP_DELTA) + Q_ARR(l, j + 1) * _
                    HW_PROBABILITIES_TREE_FUNC(-1, j + 1, JMAX_VAL, KAPPA, TEMP_DELTA) * _
                    Exp(-R_ARR(l, j + 1) * TEMP_DELTA)
        ElseIf j = ii Then
            Q_ARR(l + 1, j) = Q_ARR(l, j + 1) * HW_PROBABILITIES_TREE_FUNC(-1, j + 1, JMAX_VAL, _
                    KAPPA, TEMP_DELTA) * Exp(-R_ARR(l, j + 1) * TEMP_DELTA)
        End If
    Next j
Next l
    
For l = NSIZE To nSTEPS
    
    JPLUS_VAL = NSIZE
    JMINUS_VAL = -NSIZE
    
    TEMP_SUM = 0
    For j = -NSIZE To NSIZE
        TEMP_SUM = TEMP_SUM + Q_ARR(l, j) * Exp(-j * BTEMP_VAL * TEMP_DELTA)
    Next j
    
    ALPHA_VECTOR(l) = (Log(TEMP_SUM) - Log(HW_PRICE_FUNC(TENOR_RNG, ZERO_RNG, _
                    (l + 1) * TEMP_DELTA))) / TEMP_DELTA
    For j = JMINUS_VAL To JPLUS_VAL
        R_ARR(l, j) = ALPHA_VECTOR(l) + j * BTEMP_VAL
    Next j
    
    For j = JMINUS_VAL To JPLUS_VAL
        If JMAX_VAL = 2 Then
            Q_ARR(l + 1, 2) = Q_ARR(l, 2) * HW_PROBABILITIES_TREE_FUNC(1, 2, JMAX_VAL, KAPPA, _
                        TEMP_DELTA) * Exp(-R_ARR(l, 2) * TEMP_DELTA) + Q_ARR(l, 1) * _
                        HW_PROBABILITIES_TREE_FUNC(1, 1, JMAX_VAL, KAPPA, TEMP_DELTA) * Exp(-R_ARR(l, 1) * _
                        TEMP_DELTA)
            Q_ARR(l + 1, 1) = Q_ARR(l, 2) * HW_PROBABILITIES_TREE_FUNC(0, 2, JMAX_VAL, KAPPA, _
                        TEMP_DELTA) * Exp(-R_ARR(l, 2) * TEMP_DELTA) + Q_ARR(l, 1) * _
                        HW_PROBABILITIES_TREE_FUNC(0, 1, JMAX_VAL, KAPPA, TEMP_DELTA) * Exp(-R_ARR(l, 1) _
                        * TEMP_DELTA) + Q_ARR(l, 0) * HW_PROBABILITIES_TREE_FUNC(1, 0, JMAX_VAL, KAPPA, _
                        TEMP_DELTA) * Exp(-R_ARR(l, 0) * TEMP_DELTA)
            Q_ARR(l + 1, 0) = Q_ARR(l, 2) * HW_PROBABILITIES_TREE_FUNC(-1, 2, JMAX_VAL, KAPPA, _
                        TEMP_DELTA) * Exp(-R_ARR(l, 2) * TEMP_DELTA) + Q_ARR(l, 1) * _
                        HW_PROBABILITIES_TREE_FUNC(-1, 1, JMAX_VAL, KAPPA, TEMP_DELTA) * Exp(-R_ARR(l, 1) _
                        * TEMP_DELTA) + Q_ARR(l, 0) * HW_PROBABILITIES_TREE_FUNC(0, 0, JMAX_VAL, KAPPA, _
                        TEMP_DELTA) * Exp(-R_ARR(l, 0) * TEMP_DELTA) + Q_ARR(l, -1) * _
                        HW_PROBABILITIES_TREE_FUNC(1, -1, JMAX_VAL, KAPPA, TEMP_DELTA) * Exp(-R_ARR(l, _
                        -1) * TEMP_DELTA) + Q_ARR(l, -2) * HW_PROBABILITIES_TREE_FUNC(1, -2, JMAX_VAL, _
                        KAPPA, TEMP_DELTA) * Exp(-R_ARR(l, -2) * TEMP_DELTA)
            Q_ARR(l + 1, -1) = Q_ARR(l, 0) * HW_PROBABILITIES_TREE_FUNC(-1, 0, JMAX_VAL, KAPPA, _
                        TEMP_DELTA) * Exp(-R_ARR(l, 0) * TEMP_DELTA) + Q_ARR(l, -1) * _
                        HW_PROBABILITIES_TREE_FUNC(0, -1, JMAX_VAL, KAPPA, TEMP_DELTA) * Exp(-R_ARR(l, _
                        -1) * TEMP_DELTA) + Q_ARR(l, -2) * HW_PROBABILITIES_TREE_FUNC(1, -2, JMAX_VAL, _
                        KAPPA, TEMP_DELTA) * Exp(-R_ARR(l, -2) * TEMP_DELTA)
            Q_ARR(l + 1, -2) = Q_ARR(l, -1) * HW_PROBABILITIES_TREE_FUNC(-1, -1, JMAX_VAL, KAPPA, _
                        TEMP_DELTA) * Exp(-R_ARR(l, -1) * TEMP_DELTA) + Q_ARR(l, -2) _
                        * HW_PROBABILITIES_TREE_FUNC(-1, -2, JMAX_VAL, KAPPA, TEMP_DELTA) * _
                        Exp(-R_ARR(l, -2) * TEMP_DELTA)
        Else
            If j = JPLUS_VAL Then
                Q_ARR(l + 1, j) = Q_ARR(l, j) * HW_PROBABILITIES_TREE_FUNC(1, j, JMAX_VAL, KAPPA, _
                            TEMP_DELTA) * Exp(-R_ARR(l, j) * TEMP_DELTA) + Q_ARR(l, _
                            j - 1) * HW_PROBABILITIES_TREE_FUNC(1, j - 1, JMAX_VAL, KAPPA, TEMP_DELTA) * _
                            Exp(-R_ARR(l, j - 1) * TEMP_DELTA)
            ElseIf j = JPLUS_VAL - 1 Then
                Q_ARR(l + 1, j) = Q_ARR(l, j + 1) * HW_PROBABILITIES_TREE_FUNC(0, j + 1, JMAX_VAL, _
                            KAPPA, TEMP_DELTA) * Exp(-R_ARR(l, j + 1) * TEMP_DELTA) + _
                            Q_ARR(l, j) * HW_PROBABILITIES_TREE_FUNC(0, j, JMAX_VAL, KAPPA, TEMP_DELTA) * _
                            Exp(-R_ARR(l, j) * TEMP_DELTA) + Q_ARR(l, j - 1) * _
                            HW_PROBABILITIES_TREE_FUNC(1, j - 1, JMAX_VAL, KAPPA, TEMP_DELTA) * _
                            Exp(-R_ARR(l, j - 1) * TEMP_DELTA)
            ElseIf j = JPLUS_VAL - 2 Then
                Q_ARR(l + 1, j) = Q_ARR(l, j + 1) * HW_PROBABILITIES_TREE_FUNC(-1, j + 1, JMAX_VAL, _
                            KAPPA, TEMP_DELTA) * Exp(-R_ARR(l, j + 1) * TEMP_DELTA) + _
                            Q_ARR(l, j) * HW_PROBABILITIES_TREE_FUNC(0, j, JMAX_VAL, KAPPA, TEMP_DELTA) * _
                            Exp(-R_ARR(l, j) * TEMP_DELTA) + Q_ARR(l, j - 1) * _
                            HW_PROBABILITIES_TREE_FUNC(1, j - 1, JMAX_VAL, KAPPA, TEMP_DELTA) * _
                            Exp(-R_ARR(l, j - 1) * TEMP_DELTA) + Q_ARR(l, j + 2) * _
                            HW_PROBABILITIES_TREE_FUNC(-1, j + 2, JMAX_VAL, KAPPA, TEMP_DELTA) * _
                            Exp(-R_ARR(l, j + 2) * TEMP_DELTA)
            ElseIf j < JPLUS_VAL - 2 And j > JMINUS_VAL + 2 Then
                Q_ARR(l + 1, j) = Q_ARR(l, j + 1) * HW_PROBABILITIES_TREE_FUNC(-1, j + 1, JMAX_VAL, _
                            KAPPA, TEMP_DELTA) * Exp(-R_ARR(l, j + 1) * TEMP_DELTA) + _
                            Q_ARR(l, j) * HW_PROBABILITIES_TREE_FUNC(0, j, JMAX_VAL, KAPPA, TEMP_DELTA) * _
                            Exp(-R_ARR(l, j) * TEMP_DELTA) + Q_ARR(l, j - 1) * _
                            HW_PROBABILITIES_TREE_FUNC(1, j - 1, JMAX_VAL, KAPPA, TEMP_DELTA) * _
                            Exp(-R_ARR(l, j - 1) * TEMP_DELTA)
            ElseIf j = JMINUS_VAL + 2 Then
                Q_ARR(l + 1, j) = Q_ARR(l, j + 1) * HW_PROBABILITIES_TREE_FUNC(-1, j + 1, JMAX_VAL, _
                            KAPPA, TEMP_DELTA) * Exp(-R_ARR(l, j + 1) * TEMP_DELTA) + _
                            Q_ARR(l, j) * HW_PROBABILITIES_TREE_FUNC(0, j, JMAX_VAL, KAPPA, TEMP_DELTA) * _
                            Exp(-R_ARR(l, j) * TEMP_DELTA) + Q_ARR(l, j - 1) * _
                            HW_PROBABILITIES_TREE_FUNC(1, j - 1, JMAX_VAL, KAPPA, TEMP_DELTA) * _
                            Exp(-R_ARR(l, j - 1) * TEMP_DELTA) + Q_ARR(l, j - 2) * _
                            HW_PROBABILITIES_TREE_FUNC(1, j - 2, JMAX_VAL, KAPPA, TEMP_DELTA) * _
                            Exp(-R_ARR(l, j - 2) * TEMP_DELTA)
            ElseIf j = JMINUS_VAL + 1 Then
                Q_ARR(l + 1, j) = Q_ARR(l, j) * HW_PROBABILITIES_TREE_FUNC(0, j, JMAX_VAL, KAPPA, _
                            TEMP_DELTA) * Exp(-R_ARR(l, j) * TEMP_DELTA) + _
                            Q_ARR(l, j - 1) * HW_PROBABILITIES_TREE_FUNC(0, j - 1, JMAX_VAL, KAPPA, _
                            TEMP_DELTA) * Exp(-R_ARR(l, j - 1) * TEMP_DELTA) + _
                            Q_ARR(l, j + 1) * HW_PROBABILITIES_TREE_FUNC(-1, j + 1, JMAX_VAL, KAPPA, _
                            TEMP_DELTA) * Exp(-R_ARR(l, j + 1) * TEMP_DELTA)
            ElseIf j = JMINUS_VAL Then
                Q_ARR(l + 1, j) = Q_ARR(l, j) * HW_PROBABILITIES_TREE_FUNC(-1, j, JMAX_VAL, KAPPA, _
                            TEMP_DELTA) * Exp(-R_ARR(l, j) * TEMP_DELTA) + _
                            Q_ARR(l, j + 1) * HW_PROBABILITIES_TREE_FUNC(-1, j + 1, JMAX_VAL, KAPPA, _
                            TEMP_DELTA) * Exp(-R_ARR(l, j + 1) * TEMP_DELTA)
            End If
        End If
    Next j
Next l

For j = -MINIMUM_FUNC(nSTEPS, JMAX_VAL) To MINIMUM_FUNC(nSTEPS, JMAX_VAL)
    O_ARR(nSTEPS, j) = CASH_FLOW * MAXIMUM_FUNC(h * (Exp(TAU_VAL * _
                    (HW_ZERO_FUNC(HW_CONTINUOUS_RATE_FUNC(R_ARR(nSTEPS, j), TEMP_DELTA, FIRST_TENOR, _
                    SIGMA, KAPPA, TENOR_RNG, ZERO_RNG), FIRST_TENOR, FIRST_TENOR + _
                    DELTA_TENOR, SIGMA, KAPPA, TENOR_RNG, ZERO_RNG) - _
                    HW_ZERO_FUNC(HW_CONTINUOUS_RATE_FUNC(R_ARR(nSTEPS, j), TEMP_DELTA, _
                    FIRST_TENOR, SIGMA, KAPPA, TENOR_RNG, ZERO_RNG), FIRST_TENOR, _
                    FIRST_TENOR + TAU_VAL, SIGMA, KAPPA, TENOR_RNG, ZERO_RNG))) - _
                    Exp(TAU_VAL * (STRIKE - HW_ZERO_FUNC(HW_CONTINUOUS_RATE_FUNC(R_ARR(nSTEPS, j), _
                    TEMP_DELTA, FIRST_TENOR, SIGMA, KAPPA, TENOR_RNG, ZERO_RNG), _
                    FIRST_TENOR, FIRST_TENOR + TAU_VAL, SIGMA, KAPPA, TENOR_RNG, _
                    ZERO_RNG)))), 0)
Next j

For i = nSTEPS To NSIZE + 1 Step -1
    For j = -NSIZE To NSIZE
    If j = JMAX_VAL Then
        O_ARR(i - 1, j) = (O_ARR(i, j) * HW_PROBABILITIES_TREE_FUNC(1, j, JMAX_VAL, KAPPA, _
                    TEMP_DELTA) + O_ARR(i, j - 1) * HW_PROBABILITIES_TREE_FUNC(0, j, JMAX_VAL, KAPPA, _
                    TEMP_DELTA) + O_ARR(i, j - 2) * HW_PROBABILITIES_TREE_FUNC(-1, j, JMAX_VAL, KAPPA, _
                    TEMP_DELTA)) * Exp(-R_ARR(i - 1, j) * TEMP_DELTA)
    ElseIf j < JMAX_VAL And j > JMIN_VAL Then
        O_ARR(i - 1, j) = (O_ARR(i, j + 1) * HW_PROBABILITIES_TREE_FUNC(1, j, JMAX_VAL, KAPPA, _
                    TEMP_DELTA) + O_ARR(i, j) * HW_PROBABILITIES_TREE_FUNC(0, j, JMAX_VAL, KAPPA, _
                    TEMP_DELTA) + O_ARR(i, j - 1) * HW_PROBABILITIES_TREE_FUNC(-1, j, JMAX_VAL, KAPPA, _
                    TEMP_DELTA)) * Exp(-R_ARR(i - 1, j) * TEMP_DELTA)
    ElseIf j = JMIN_VAL Then
        O_ARR(i - 1, j) = (O_ARR(i, j) * HW_PROBABILITIES_TREE_FUNC(-1, j, JMAX_VAL, KAPPA, _
                    TEMP_DELTA) + O_ARR(i, j + 1) * HW_PROBABILITIES_TREE_FUNC(0, j, JMAX_VAL, KAPPA, _
                    TEMP_DELTA) + O_ARR(i, j + 2) * HW_PROBABILITIES_TREE_FUNC(1, j, JMAX_VAL, KAPPA, _
                    TEMP_DELTA)) * Exp(-R_ARR(i - 1, j) * TEMP_DELTA)
    End If
    Next j
Next i

JMINUS_VAL = -NSIZE
JPLUS_VAL = NSIZE

For i = NSIZE To 1 Step -1
    JMINUS_VAL = JMINUS_VAL + 1
    JPLUS_VAL = JPLUS_VAL - 1
    For j = JMINUS_VAL To JPLUS_VAL
         O_ARR(i - 1, j) = (O_ARR(i, j + 1) * HW_PROBABILITIES_TREE_FUNC(1, j, JMAX_VAL, KAPPA, _
                TEMP_DELTA) + O_ARR(i, j) * HW_PROBABILITIES_TREE_FUNC(0, j, JMAX_VAL, KAPPA, _
                TEMP_DELTA) + O_ARR(i, j - 1) * HW_PROBABILITIES_TREE_FUNC(-1, j, JMAX_VAL, KAPPA, _
                TEMP_DELTA)) * Exp(-R_ARR(i - 1, j) * TEMP_DELTA)
    Next j
Next i


HW_CAP_FLOOR_LET_FUNC = O_ARR(0, 0)
   
Exit Function
ERROR_LABEL:
    HW_CAP_FLOOR_LET_FUNC = Err.number
End Function


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : HW_SWAPTION_FUNC
'DESCRIPTION   : 'Options on interest swaps or swaptions as they are also called
'are another popular instrument to hedge against unfavorable interest rate movements.
'They give the holder the possibility to enter into a specified interest
'rate swap at some future time. As this is only a right but no obligation
'it must have some value.

'To value a swaption we make use of the fact that a swap can be regarded as
'the agreement to exchange a fixed rate bond for a floating-rate bond. As
'the value of the floating rate bond always equals the principal amount at
'initiation of the swap regardless of the specific floating rate it is based
'on, we can say that a swaption is an option to exchange the fixed rate bond
'for the principal amount of the swap. This means that for an option on a
'payer swap where we pay fixed and receive floating the swaption can be
'replicated by a put option on the fixed rate bond with strike price equal to
'the principal. A receiver swaption is valued similar but with a call option.

'LIBRARY       : FIXED_INCOME
'GROUP         : HW_1F
'ID            : 008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'REFERENCE: http://www.angelfire.com/ny/financeinfo/HullWhite.pdf
'**********************************************************************************
'**********************************************************************************


Function HW_SWAPTION_FUNC(ByVal OPTION_FLAG As Integer, _
ByVal TEMP_DELTA As Double, _
ByVal SIGMA As Double, _
ByVal KAPPA As Double, _
ByVal FIRST_TENOR_RNG As Variant, _
ByVal FIRST_TENOR As Double, _
ByVal CASH_FLOW As Double, _
ByVal STRIKE_RATE As Double, _
ByRef TENOR_RNG As Variant, _
ByRef ZERO_RNG As Variant)

Dim h As Double
Dim i As Double
Dim j As Double
Dim k As Double
Dim l As Double

Dim ii As Double
Dim jj As Double

Dim JMAX_VAL As Double
Dim JMIN_VAL As Double

Dim JPLUS_VAL As Double
Dim JMINUS_VAL As Double

Dim NSIZE As Double
Dim NROWS As Double
Dim nSTEPS As Double

Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double
Dim CTEMP_VAL As Double

Dim TEMP_SUM As Double

Dim Q_ARR As Variant
Dim O_ARR As Variant
Dim R_ARR As Variant
Dim S_ARR As Variant
Dim T_ARR As Variant

Dim ALPHA_VECTOR As Variant
Dim STEPS_VECTOR As Variant
Dim FIRST_TENOR_VECTOR As Variant

On Error GoTo ERROR_LABEL

If IsArray(FIRST_TENOR_RNG) = True Then
    FIRST_TENOR_VECTOR = FIRST_TENOR_RNG
    If UBound(FIRST_TENOR_VECTOR, 1) = 1 Then: _
        FIRST_TENOR_VECTOR = MATRIX_TRANSPOSE_FUNC(FIRST_TENOR_VECTOR)
Else
    ReDim FIRST_TENOR_VECTOR(1 To 1, 1 To 1)
    FIRST_TENOR_VECTOR(1, 1) = FIRST_TENOR_RNG
End If

NROWS = UBound(FIRST_TENOR_VECTOR, 1)
Select Case OPTION_FLAG
    Case -1 '"p", "put", -1
        h = 1
    Case Else
        h = -1
End Select

nSTEPS = FIRST_TENOR / TEMP_DELTA

'calculating constants
ATEMP_VAL = SIGMA ^ 2 * (1 - Exp(-2 * KAPPA * TEMP_DELTA)) / (2 * KAPPA)
BTEMP_VAL = Sqr(3 * ATEMP_VAL)
JMAX_VAL = Int(0.184 / (KAPPA * TEMP_DELTA)) + 1
JMIN_VAL = -JMAX_VAL
NSIZE = MINIMUM_FUNC(nSTEPS, JMAX_VAL)

'redimensioning variables
ReDim ALPHA_VECTOR(0 To nSTEPS)
ReDim Q_ARR(0 To nSTEPS + 1, JMIN_VAL To JMAX_VAL)
ReDim O_ARR(0 To nSTEPS, JMIN_VAL To JMAX_VAL)
ReDim R_ARR(0 To nSTEPS, JMIN_VAL To JMAX_VAL)
ReDim S_ARR(0 To nSTEPS, JMIN_VAL To JMAX_VAL)
ReDim T_ARR(0 To nSTEPS, JMIN_VAL To JMAX_VAL)
ReDim STEPS_VECTOR(0 To NROWS)

'calculating start values
ALPHA_VECTOR(0) = -Log(HW_PRICE_FUNC(TENOR_RNG, ZERO_RNG, TEMP_DELTA)) / TEMP_DELTA
R_ARR(0, 0) = ALPHA_VECTOR(0)
Q_ARR(1, 1) = HW_PROBABILITIES_TREE_FUNC(1, 0, JMAX_VAL, KAPPA, TEMP_DELTA) * Exp(-R_ARR(0, 0) * TEMP_DELTA)
Q_ARR(1, 0) = HW_PROBABILITIES_TREE_FUNC(0, 0, JMAX_VAL, KAPPA, TEMP_DELTA) * Exp(-R_ARR(0, 0) * TEMP_DELTA)
Q_ARR(1, -1) = HW_PROBABILITIES_TREE_FUNC(-1, 0, JMAX_VAL, KAPPA, TEMP_DELTA) * Exp(-R_ARR(0, 0) * TEMP_DELTA)

JMINUS_VAL = 0
JPLUS_VAL = 0

ii = -1
jj = 1

'using Q_ARR(l) to calculate ALPHA_VECTOR(l),R_ARR(l) and then Q_ARR(m+1) up to
'NSIZE-1 so we get the Q_ARR(NSIZE) as last results
For l = 1 To NSIZE - 1
    JMINUS_VAL = JMINUS_VAL - 1
    JPLUS_VAL = JPLUS_VAL + 1
    TEMP_SUM = 0
    For j = JMINUS_VAL To JPLUS_VAL
        TEMP_SUM = TEMP_SUM + Q_ARR(l, j) * Exp(-j * BTEMP_VAL * TEMP_DELTA)
    Next j
    
    ALPHA_VECTOR(l) = (Log(TEMP_SUM) - Log(HW_PRICE_FUNC(TENOR_RNG, ZERO_RNG, _
                (l + 1) * TEMP_DELTA))) / TEMP_DELTA
    For j = JMINUS_VAL To JPLUS_VAL
        R_ARR(l, j) = ALPHA_VECTOR(l) + j * BTEMP_VAL
    Next j
    
    ii = ii - 1
    jj = jj + 1
    
    For j = ii To jj
    
        If j = jj Then
            Q_ARR(l + 1, j) = Q_ARR(l, j - 1) * HW_PROBABILITIES_TREE_FUNC(1, j - 1, JMAX_VAL, _
                        KAPPA, TEMP_DELTA) * Exp(-R_ARR(l, j - 1) * TEMP_DELTA)
        ElseIf j = jj - 1 Then
            Q_ARR(l + 1, j) = Q_ARR(l, j) * HW_PROBABILITIES_TREE_FUNC(0, j, JMAX_VAL, KAPPA, _
                        TEMP_DELTA) * Exp(-R_ARR(l, j) * TEMP_DELTA) + Q_ARR(l, j - 1) _
                        * HW_PROBABILITIES_TREE_FUNC(1, j - 1, JMAX_VAL, KAPPA, TEMP_DELTA) * _
                        Exp(-R_ARR(l, j - 1) * TEMP_DELTA)
        ElseIf j < jj - 1 And j > ii + 1 Then
            Q_ARR(l + 1, j) = Q_ARR(l, j - 1) * HW_PROBABILITIES_TREE_FUNC(1, j - 1, JMAX_VAL, _
                        KAPPA, TEMP_DELTA) * Exp(-R_ARR(l, j - 1) * TEMP_DELTA) + _
                        Q_ARR(l, j) * HW_PROBABILITIES_TREE_FUNC(0, j, JMAX_VAL, KAPPA, TEMP_DELTA) * _
                        Exp(-R_ARR(l, j) * TEMP_DELTA) + Q_ARR(l, j + 1) * _
                        HW_PROBABILITIES_TREE_FUNC(-1, j + 1, JMAX_VAL, KAPPA, TEMP_DELTA) * _
                        Exp(-R_ARR(l, j + 1) * TEMP_DELTA)
        ElseIf j = ii + 1 Then
            Q_ARR(l + 1, j) = Q_ARR(l, j) * HW_PROBABILITIES_TREE_FUNC(0, j, JMAX_VAL, KAPPA, _
                        TEMP_DELTA) * Exp(-R_ARR(l, j) * TEMP_DELTA) + Q_ARR(l, _
                        j + 1) * HW_PROBABILITIES_TREE_FUNC(-1, j + 1, JMAX_VAL, KAPPA, TEMP_DELTA) * _
                        Exp(-R_ARR(l, j + 1) * TEMP_DELTA)
        ElseIf j = ii Then
            Q_ARR(l + 1, j) = Q_ARR(l, j + 1) * HW_PROBABILITIES_TREE_FUNC(-1, j + 1, JMAX_VAL, KAPPA, _
                        TEMP_DELTA) * Exp(-R_ARR(l, j + 1) * TEMP_DELTA)
        End If
    Next j
Next l
    
    'using the Q_ARR(NSIZE) to calculate ALPHA_VECTOR(NSIZE) and so on up to nSTEPS
For l = NSIZE To nSTEPS
    
    JPLUS_VAL = JMAX_VAL
    JMINUS_VAL = JMIN_VAL
    
    TEMP_SUM = 0
    For j = JMIN_VAL To JMAX_VAL
        TEMP_SUM = TEMP_SUM + Q_ARR(l, j) * Exp(-j * BTEMP_VAL * TEMP_DELTA)
    Next j
    
    ALPHA_VECTOR(l) = (Log(TEMP_SUM) - Log(HW_PRICE_FUNC(TENOR_RNG, ZERO_RNG, _
            (l + 1) * TEMP_DELTA))) / TEMP_DELTA
    For j = JMINUS_VAL To JPLUS_VAL
        R_ARR(l, j) = ALPHA_VECTOR(l) + j * BTEMP_VAL
    Next j
    
    For j = JMINUS_VAL To JPLUS_VAL
    'considering the spezial case when nonstandard branching leads to central nodes
    'with five incoming arrows
        If JMAX_VAL = 2 Then
            Q_ARR(l + 1, 2) = Q_ARR(l, 2) * HW_PROBABILITIES_TREE_FUNC(1, 2, JMAX_VAL, KAPPA, _
                        TEMP_DELTA) * Exp(-R_ARR(l, 2) * TEMP_DELTA) + Q_ARR(l, 1) * _
                        HW_PROBABILITIES_TREE_FUNC(1, 1, JMAX_VAL, KAPPA, TEMP_DELTA) * Exp(-R_ARR(l, 1) _
                        * TEMP_DELTA)
            Q_ARR(l + 1, 1) = Q_ARR(l, 2) * HW_PROBABILITIES_TREE_FUNC(0, 2, JMAX_VAL, KAPPA, _
                        TEMP_DELTA) * Exp(-R_ARR(l, 2) * TEMP_DELTA) + Q_ARR(l, 1) * _
                        HW_PROBABILITIES_TREE_FUNC(0, 1, JMAX_VAL, KAPPA, TEMP_DELTA) * Exp(-R_ARR(l, 1) _
                        * TEMP_DELTA) + Q_ARR(l, 0) * HW_PROBABILITIES_TREE_FUNC(1, 0, JMAX_VAL, KAPPA, _
                        TEMP_DELTA) * Exp(-R_ARR(l, 0) * TEMP_DELTA)
            Q_ARR(l + 1, 0) = Q_ARR(l, 2) * HW_PROBABILITIES_TREE_FUNC(-1, 2, JMAX_VAL, KAPPA, _
                        TEMP_DELTA) * Exp(-R_ARR(l, 2) * TEMP_DELTA) + Q_ARR(l, 1) * _
                        HW_PROBABILITIES_TREE_FUNC(-1, 1, JMAX_VAL, KAPPA, TEMP_DELTA) * Exp(-R_ARR(l, 1) _
                        * TEMP_DELTA) + Q_ARR(l, 0) * HW_PROBABILITIES_TREE_FUNC(0, 0, JMAX_VAL, KAPPA, _
                        TEMP_DELTA) * Exp(-R_ARR(l, 0) * TEMP_DELTA) + Q_ARR(l, -1) * _
                        HW_PROBABILITIES_TREE_FUNC(1, -1, JMAX_VAL, KAPPA, TEMP_DELTA) * Exp(-R_ARR(l, -1) _
                        * TEMP_DELTA) + Q_ARR(l, -2) * HW_PROBABILITIES_TREE_FUNC(1, -2, JMAX_VAL, KAPPA, _
                        TEMP_DELTA) * Exp(-R_ARR(l, -2) * TEMP_DELTA)
            Q_ARR(l + 1, -1) = Q_ARR(l, 0) * HW_PROBABILITIES_TREE_FUNC(-1, 0, JMAX_VAL, KAPPA, _
                        TEMP_DELTA) * Exp(-R_ARR(l, 0) * TEMP_DELTA) + Q_ARR(l, -1) * _
                        HW_PROBABILITIES_TREE_FUNC(0, -1, JMAX_VAL, KAPPA, TEMP_DELTA) * Exp(-R_ARR(l, -1) _
                        * TEMP_DELTA) + Q_ARR(l, -2) * HW_PROBABILITIES_TREE_FUNC(1, -2, JMAX_VAL, KAPPA, _
                        TEMP_DELTA) * Exp(-R_ARR(l, -2) * TEMP_DELTA)
            Q_ARR(l + 1, -2) = Q_ARR(l, -1) * HW_PROBABILITIES_TREE_FUNC(-1, -1, JMAX_VAL, KAPPA, _
                        TEMP_DELTA) * Exp(-R_ARR(l, -1) * TEMP_DELTA) + Q_ARR(l, -2) * _
                        HW_PROBABILITIES_TREE_FUNC(-1, -2, JMAX_VAL, KAPPA, TEMP_DELTA) * Exp(-R_ARR(l, -2) _
                        * TEMP_DELTA)
        Else
            If j = JPLUS_VAL Then
                Q_ARR(l + 1, j) = Q_ARR(l, j) * HW_PROBABILITIES_TREE_FUNC(1, j, JMAX_VAL, KAPPA, _
                            TEMP_DELTA) * Exp(-R_ARR(l, j) * TEMP_DELTA) + Q_ARR(l, _
                            j - 1) * HW_PROBABILITIES_TREE_FUNC(1, j - 1, JMAX_VAL, KAPPA, TEMP_DELTA) * _
                            Exp(-R_ARR(l, j - 1) * TEMP_DELTA)
            ElseIf j = JPLUS_VAL - 1 Then
                Q_ARR(l + 1, j) = Q_ARR(l, j + 1) * HW_PROBABILITIES_TREE_FUNC(0, j + 1, JMAX_VAL, _
                            KAPPA, TEMP_DELTA) * Exp(-R_ARR(l, j + 1) * TEMP_DELTA) + _
                            Q_ARR(l, j) * HW_PROBABILITIES_TREE_FUNC(0, j, JMAX_VAL, KAPPA, TEMP_DELTA) * _
                            Exp(-R_ARR(l, j) * TEMP_DELTA) + Q_ARR(l, j - 1) * _
                            HW_PROBABILITIES_TREE_FUNC(1, j - 1, JMAX_VAL, KAPPA, TEMP_DELTA) * _
                            Exp(-R_ARR(l, j - 1) * TEMP_DELTA)
            ElseIf j = JPLUS_VAL - 2 Then
                Q_ARR(l + 1, j) = Q_ARR(l, j + 1) * HW_PROBABILITIES_TREE_FUNC(-1, j + 1, JMAX_VAL, _
                            KAPPA, TEMP_DELTA) * Exp(-R_ARR(l, j + 1) * TEMP_DELTA) + _
                            Q_ARR(l, j) * HW_PROBABILITIES_TREE_FUNC(0, j, JMAX_VAL, KAPPA, TEMP_DELTA) * _
                            Exp(-R_ARR(l, j) * TEMP_DELTA) + Q_ARR(l, j - 1) * _
                            HW_PROBABILITIES_TREE_FUNC(1, j - 1, JMAX_VAL, KAPPA, TEMP_DELTA) * _
                            Exp(-R_ARR(l, j - 1) * TEMP_DELTA) + Q_ARR(l, j + 2) * _
                            HW_PROBABILITIES_TREE_FUNC(-1, j + 2, JMAX_VAL, KAPPA, TEMP_DELTA) * _
                            Exp(-R_ARR(l, j + 2) * TEMP_DELTA)
            ElseIf j < JPLUS_VAL - 2 And j > JMINUS_VAL + 2 Then
                Q_ARR(l + 1, j) = Q_ARR(l, j + 1) * HW_PROBABILITIES_TREE_FUNC(-1, j + 1, JMAX_VAL, _
                            KAPPA, TEMP_DELTA) * Exp(-R_ARR(l, j + 1) * TEMP_DELTA) + _
                            Q_ARR(l, j) * HW_PROBABILITIES_TREE_FUNC(0, j, JMAX_VAL, KAPPA, TEMP_DELTA) * _
                            Exp(-R_ARR(l, j) * TEMP_DELTA) + Q_ARR(l, j - 1) * _
                            HW_PROBABILITIES_TREE_FUNC(1, j - 1, JMAX_VAL, KAPPA, TEMP_DELTA) * _
                            Exp(-R_ARR(l, j - 1) * TEMP_DELTA)
            ElseIf j = JMINUS_VAL + 2 Then
                Q_ARR(l + 1, j) = Q_ARR(l, j + 1) * HW_PROBABILITIES_TREE_FUNC(-1, j + 1, JMAX_VAL, _
                            KAPPA, TEMP_DELTA) * Exp(-R_ARR(l, j + 1) * TEMP_DELTA) + _
                            Q_ARR(l, j) * HW_PROBABILITIES_TREE_FUNC(0, j, JMAX_VAL, KAPPA, TEMP_DELTA) * _
                            Exp(-R_ARR(l, j) * TEMP_DELTA) + Q_ARR(l, j - 1) * _
                            HW_PROBABILITIES_TREE_FUNC(1, j - 1, JMAX_VAL, KAPPA, TEMP_DELTA) * _
                            Exp(-R_ARR(l, j - 1) * TEMP_DELTA) + Q_ARR(l, j - 2) * _
                            HW_PROBABILITIES_TREE_FUNC(1, j - 2, JMAX_VAL, KAPPA, TEMP_DELTA) * _
                            Exp(-R_ARR(l, j - 2) * TEMP_DELTA)
            ElseIf j = JMINUS_VAL + 1 Then
                Q_ARR(l + 1, j) = Q_ARR(l, j) * HW_PROBABILITIES_TREE_FUNC(0, j, JMAX_VAL, KAPPA, _
                            TEMP_DELTA) * Exp(-R_ARR(l, j) * TEMP_DELTA) + _
                            Q_ARR(l, j - 1) * HW_PROBABILITIES_TREE_FUNC(0, j - 1, JMAX_VAL, _
                            KAPPA, TEMP_DELTA) * Exp(-R_ARR(l, j - 1) * _
                            TEMP_DELTA) + Q_ARR(l, j + 1) * HW_PROBABILITIES_TREE_FUNC(-1, _
                            j + 1, JMAX_VAL, KAPPA, TEMP_DELTA) * Exp(-R_ARR(l, _
                            j + 1) * TEMP_DELTA)
            ElseIf j = JMINUS_VAL Then
                Q_ARR(l + 1, j) = Q_ARR(l, j) * HW_PROBABILITIES_TREE_FUNC(-1, j, JMAX_VAL, KAPPA, _
                            TEMP_DELTA) * Exp(-R_ARR(l, j) * TEMP_DELTA) + _
                            Q_ARR(l, j + 1) * HW_PROBABILITIES_TREE_FUNC(-1, j + 1, JMAX_VAL, _
                            KAPPA, TEMP_DELTA) * Exp(-R_ARR(l, j + 1) * TEMP_DELTA)
            End If
        End If
    Next j
Next l

'cash flow calculations
    CTEMP_VAL = FIRST_TENOR_VECTOR(2, 1) - FIRST_TENOR_VECTOR(1, 1)
    For j = -MINIMUM_FUNC(nSTEPS, JMAX_VAL) To MINIMUM_FUNC(nSTEPS, JMAX_VAL)
        S_ARR(nSTEPS, j) = HW_SWAP_SOLVER_FUNC(0, R_ARR(nSTEPS, j), KAPPA, TEMP_DELTA, _
                SIGMA, TENOR_RNG, ZERO_RNG, FIRST_TENOR, FIRST_TENOR_VECTOR)
    Next j
     
For k = 1 To NROWS
    For j = -MINIMUM_FUNC(nSTEPS, JMAX_VAL) To MINIMUM_FUNC(nSTEPS, JMAX_VAL)
        T_ARR(nSTEPS, j) = HW_ZERO_FUNC(HW_CONTINUOUS_RATE_FUNC(R_ARR(nSTEPS, j), _
                    TEMP_DELTA, FIRST_TENOR, SIGMA, KAPPA, TENOR_RNG, ZERO_RNG), _
                    FIRST_TENOR, FIRST_TENOR_VECTOR(k, 1), SIGMA, KAPPA, TENOR_RNG, _
                    ZERO_RNG)
        O_ARR(nSTEPS, j) = CASH_FLOW * MAXIMUM_FUNC(h * (Exp(CTEMP_VAL * S_ARR(nSTEPS, j)) _
                    - Exp(CTEMP_VAL * STRIKE_RATE)), 0) * Exp(-T_ARR(nSTEPS, j) * _
                    (FIRST_TENOR_VECTOR(k, 1) - FIRST_TENOR)) + O_ARR(nSTEPS, j)
    Next j
Next k

'backward induction
For i = nSTEPS To NSIZE + 1 Step -1
    For j = JMIN_VAL To JMAX_VAL
        If j = JMAX_VAL Then
            O_ARR(i - 1, j) = (O_ARR(i, j) * HW_PROBABILITIES_TREE_FUNC(1, j, JMAX_VAL, KAPPA, _
                        TEMP_DELTA) + O_ARR(i, j - 1) * HW_PROBABILITIES_TREE_FUNC(0, j, JMAX_VAL, _
                        KAPPA, TEMP_DELTA) + O_ARR(i, j - 2) * HW_PROBABILITIES_TREE_FUNC(-1, j, _
                        JMAX_VAL, KAPPA, TEMP_DELTA)) * Exp(-R_ARR(i - 1, j) * _
                        TEMP_DELTA) + O_ARR(i - 1, j)
        ElseIf j < JMAX_VAL And j > JMIN_VAL Then
            O_ARR(i - 1, j) = (O_ARR(i, j + 1) * HW_PROBABILITIES_TREE_FUNC(1, j, JMAX_VAL, KAPPA, _
                    TEMP_DELTA) + O_ARR(i, j) * HW_PROBABILITIES_TREE_FUNC(0, j, JMAX_VAL, KAPPA, _
                    TEMP_DELTA) + O_ARR(i, j - 1) * HW_PROBABILITIES_TREE_FUNC(-1, j, JMAX_VAL, KAPPA, _
                    TEMP_DELTA)) * Exp(-R_ARR(i - 1, j) * TEMP_DELTA) + O_ARR(i - 1, j)
        ElseIf j = JMIN_VAL Then
            O_ARR(i - 1, j) = (O_ARR(i, j) * HW_PROBABILITIES_TREE_FUNC(-1, j, JMAX_VAL, KAPPA, _
                    TEMP_DELTA) + O_ARR(i, j + 1) * HW_PROBABILITIES_TREE_FUNC(0, j, JMAX_VAL, KAPPA, _
                    TEMP_DELTA) + O_ARR(i, j + 2) * HW_PROBABILITIES_TREE_FUNC(1, j, JMAX_VAL, KAPPA, _
                    TEMP_DELTA)) * Exp(-R_ARR(i - 1, j) * TEMP_DELTA) + O_ARR(i - 1, j)
        End If
    Next j
Next i

JMINUS_VAL = -NSIZE
JPLUS_VAL = NSIZE

For i = NSIZE To 1 Step -1
    JMINUS_VAL = JMINUS_VAL + 1
    JPLUS_VAL = JPLUS_VAL - 1
    For j = JMINUS_VAL To JPLUS_VAL
         O_ARR(i - 1, j) = (O_ARR(i, j + 1) * HW_PROBABILITIES_TREE_FUNC(1, j, JMAX_VAL, _
                        KAPPA, TEMP_DELTA) + O_ARR(i, j) * HW_PROBABILITIES_TREE_FUNC(0, j, _
                        JMAX_VAL, KAPPA, TEMP_DELTA) + O_ARR(i, j - 1) * _
                        HW_PROBABILITIES_TREE_FUNC(-1, j, JMAX_VAL, KAPPA, TEMP_DELTA)) * _
                        Exp(-R_ARR(i - 1, j) * TEMP_DELTA) + O_ARR(i - 1, j)
     Next j
Next i


HW_SWAPTION_FUNC = O_ARR(0, 0)
   
Exit Function
ERROR_LABEL:
HW_SWAPTION_FUNC = Err.number
End Function


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : HW_SWAP_OBJ_FUNC
'DESCRIPTION   : HW Swap Solver Obj Function
'LIBRARY       : FIXED_INCOME
'GROUP         : HW_1F
'ID            : 009
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'REFERENCE: http://www.angelfire.com/ny/financeinfo/HullWhite.pdf
'**********************************************************************************
'**********************************************************************************


Function HW_SWAP_OBJ_FUNC(ByVal ZERO_RATE As Double, _
ByVal DISC_RATE As Double, _
ByVal KAPPA As Double, _
ByVal DELTA_TENOR As Double, _
ByVal SIGMA As Double, _
ByRef TENOR_RNG As Variant, _
ByRef ZERO_RNG As Variant, _
ByVal FIRST_TENOR As Double, _
ByRef FIRST_TENOR_RNG As Variant)

Dim i As Double
Dim j As Double

Dim NSIZE As Double
Dim BTEMP_VAL As Double

Dim TEMP_DIFF As Double
Dim TEMP_SUM As Double
Dim FIRST_TENOR_VECTOR As Variant

On Error GoTo ERROR_LABEL

If IsArray(FIRST_TENOR_RNG) = True Then
    FIRST_TENOR_VECTOR = FIRST_TENOR_RNG
    If UBound(FIRST_TENOR_VECTOR, 1) = 1 Then: _
        FIRST_TENOR_VECTOR = MATRIX_TRANSPOSE_FUNC(FIRST_TENOR_VECTOR)
Else
    ReDim FIRST_TENOR_VECTOR(1 To 1, 1 To 1)
    FIRST_TENOR_VECTOR(1, 1) = FIRST_TENOR_RNG
End If

BTEMP_VAL = ZERO_RATE
NSIZE = UBound(FIRST_TENOR_VECTOR, 1)
TEMP_DIFF = FIRST_TENOR_VECTOR(2, 1) - FIRST_TENOR_VECTOR(1, 1)

TEMP_SUM = 0
For i = 1 To NSIZE
    If i = NSIZE Then j = 0 Else j = 1
    TEMP_SUM = TEMP_SUM + 100 * (Exp(BTEMP_VAL * TEMP_DIFF) - j) * _
        Exp(-HW_ZERO_FUNC(HW_CONTINUOUS_RATE_FUNC(DISC_RATE, DELTA_TENOR, _
        FIRST_TENOR, SIGMA, KAPPA, TENOR_RNG, ZERO_RNG), FIRST_TENOR, _
        FIRST_TENOR_VECTOR(i, 1), SIGMA, KAPPA, _
        TENOR_RNG, ZERO_RNG) * (FIRST_TENOR_VECTOR(i, 1) - FIRST_TENOR))

Next i

HW_SWAP_OBJ_FUNC = TEMP_SUM - 100

Exit Function
ERROR_LABEL:
    HW_SWAP_OBJ_FUNC = Err.number
End Function


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : HW_SWAP_SOLVER_FUNC
'DESCRIPTION   : HP Swap Newton Solver Function
'LIBRARY       : FIXED_INCOME
'GROUP         : HW_1F
'ID            : 010
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'REFERENCE: http://www.angelfire.com/ny/financeinfo/HullWhite.pdf
'**********************************************************************************
'**********************************************************************************

Private Function HW_SWAP_SOLVER_FUNC(ByVal ZERO_RATE As Double, _
ByVal DISC_RATE As Double, _
ByVal KAPPA As Double, _
ByVal DELTA_TENOR As Double, _
ByVal SIGMA As Double, _
ByRef TENOR_RNG As Variant, _
ByRef ZERO_RNG As Variant, _
ByVal FIRST_TENOR As Double, _
ByRef FIRST_TENOR_RNG As Variant)

Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double
Dim CTEMP_VAL As Double

Dim epsilon As Double
Dim tolerance As Double

On Error GoTo ERROR_LABEL

epsilon = 0.01
tolerance = 0.000001

BTEMP_VAL = ZERO_RATE

Do
    ATEMP_VAL = (HW_SWAP_OBJ_FUNC(BTEMP_VAL + epsilon, DISC_RATE, KAPPA, DELTA_TENOR, _
                SIGMA, TENOR_RNG, ZERO_RNG, FIRST_TENOR, FIRST_TENOR_RNG) _
                - HW_SWAP_OBJ_FUNC(BTEMP_VAL - epsilon, DISC_RATE, KAPPA, DELTA_TENOR, _
                SIGMA, TENOR_RNG, ZERO_RNG, FIRST_TENOR, FIRST_TENOR_RNG)) / (2 * epsilon)
    BTEMP_VAL = BTEMP_VAL - HW_SWAP_OBJ_FUNC(BTEMP_VAL, DISC_RATE, KAPPA, DELTA_TENOR, SIGMA, _
                TENOR_RNG, ZERO_RNG, FIRST_TENOR, FIRST_TENOR_RNG) / ATEMP_VAL
    CTEMP_VAL = Abs(HW_SWAP_OBJ_FUNC(BTEMP_VAL, DISC_RATE, KAPPA, DELTA_TENOR, SIGMA, _
                TENOR_RNG, ZERO_RNG, FIRST_TENOR, FIRST_TENOR_RNG) - 0)
Loop Until CTEMP_VAL <= tolerance

HW_SWAP_SOLVER_FUNC = BTEMP_VAL

Exit Function
ERROR_LABEL:
    HW_SWAP_SOLVER_FUNC = Err.number
End Function

'**********************************************************************************
'**********************************************************************************
'FUNCTION      : HW_BINARY_FUNC

'DESCRIPTION   : Value of binary options (Accrual Swaps)
'The numerical procedure presented above can also be used to price accrual
'swaps. These are swaps where the interest on one side accrues only if the
'floating rate is in a certain range or above or below a rate.

'Like Hull (1997) points out an accrual swap can be replicated by an ordinary
'swap and a series of binary options. For every day f and state j of the swap
'we have to price a binary option which provides a payoff at the following
'swap payment date. If for example the floating reference rate is below the
'strike rate RcX. RD is the discount rate for the period between calculating
'the cash flow at time f*dt and the next swap payment date s(f) and tau
'the number of days between swap payments assuming the year has 248 business
'days

'LIBRARY       : FIXED_INCOME
'GROUP         : HW_1F
'ID            : 011
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'REFERENCE: http://www.angelfire.com/ny/financeinfo/HullWhite.pdf
'**********************************************************************************
'**********************************************************************************

Function HW_BINARY_FUNC(ByVal SIGMA As Double, _
ByVal KAPPA As Double, _
ByVal FIRST_TENOR As Double, _
ByVal TAU_VAL As Double, _
ByVal CASH_FLOW As Double, _
ByVal STRIKE_RATE As Double, _
ByRef TENOR_RNG As Variant, _
ByRef ZERO_RNG As Variant, _
Optional ByVal TRADING_DAYS As Double = 248)

Dim i As Double
Dim j As Double
Dim k As Double
Dim l As Double

Dim ii As Double
Dim jj As Double

Dim JMAX_VAL As Double
Dim JMIN_VAL As Double

Dim JPLUS_VAL As Double
Dim JMINUS_VAL As Double

Dim NSIZE As Double
Dim nSTEPS As Double

Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double

Dim TEMP_SUM As Double
Dim TEMP_DELTA As Double

Dim Q_ARR As Variant
Dim O_ARR As Variant
Dim R_ARR As Variant
Dim S_VECTOR As Variant
Dim S_MATRIX As Variant

Dim ALPHA_VECTOR As Variant

On Error GoTo ERROR_LABEL

TEMP_DELTA = 1 / TRADING_DAYS  'step is one business day

nSTEPS = FIRST_TENOR / TEMP_DELTA

'calculating constants
ATEMP_VAL = SIGMA ^ 2 * (1 - Exp(-2 * KAPPA * TEMP_DELTA)) / (2 * KAPPA)
BTEMP_VAL = Sqr(3 * ATEMP_VAL)
JMAX_VAL = Int(0.184 / (KAPPA * TEMP_DELTA)) + 1
JMIN_VAL = -JMAX_VAL
NSIZE = MINIMUM_FUNC(nSTEPS, JMAX_VAL)

'redimensioning variables
ReDim ALPHA_VECTOR(0 To nSTEPS)
ReDim Q_ARR(0 To nSTEPS + 1, -NSIZE To NSIZE)
ReDim O_ARR(0 To nSTEPS, -NSIZE To NSIZE)
ReDim R_ARR(0 To nSTEPS, -NSIZE To NSIZE)
ReDim S_MATRIX(0 To nSTEPS, -NSIZE To NSIZE)
ReDim S_VECTOR(0 To nSTEPS)

'calculating start values
ALPHA_VECTOR(0) = -Log(HW_PRICE_FUNC(TENOR_RNG, ZERO_RNG, TEMP_DELTA)) / TEMP_DELTA
R_ARR(0, 0) = ALPHA_VECTOR(0)
S_MATRIX(0, 0) = HW_CONTINUOUS_RATE_FUNC(R_ARR(0, 0), TEMP_DELTA, 0, _
                SIGMA, KAPPA, TENOR_RNG, ZERO_RNG)
Q_ARR(1, 1) = HW_PROBABILITIES_TREE_FUNC(1, 0, JMAX_VAL, KAPPA, TEMP_DELTA) * _
            Exp(-R_ARR(0, 0) * TEMP_DELTA)
Q_ARR(1, 0) = HW_PROBABILITIES_TREE_FUNC(0, 0, JMAX_VAL, KAPPA, TEMP_DELTA) * _
                Exp(-R_ARR(0, 0) * TEMP_DELTA)
Q_ARR(1, -1) = HW_PROBABILITIES_TREE_FUNC(-1, 0, JMAX_VAL, KAPPA, TEMP_DELTA) * _
            Exp(-R_ARR(0, 0) * TEMP_DELTA)

JMINUS_VAL = 0
JPLUS_VAL = 0

ii = -1
jj = 1

For l = 1 To NSIZE - 1
    JMINUS_VAL = JMINUS_VAL - 1
    JPLUS_VAL = JPLUS_VAL + 1
    TEMP_SUM = 0
    For j = JMINUS_VAL To JPLUS_VAL
        TEMP_SUM = TEMP_SUM + Q_ARR(l, j) * Exp(-j * BTEMP_VAL * TEMP_DELTA)
    Next j
    
    ALPHA_VECTOR(l) = (Log(TEMP_SUM) - Log(HW_PRICE_FUNC(TENOR_RNG, ZERO_RNG, _
                    (l + 1) * TEMP_DELTA))) / TEMP_DELTA
    For j = JMINUS_VAL To JPLUS_VAL
        R_ARR(l, j) = ALPHA_VECTOR(l) + j * BTEMP_VAL
        S_MATRIX(l, j) = HW_CONTINUOUS_RATE_FUNC(R_ARR(l, j), TEMP_DELTA, l * _
                            TEMP_DELTA, SIGMA, KAPPA, TENOR_RNG, ZERO_RNG)
    Next j
    
    ii = ii - 1
    jj = jj + 1
    
    For j = ii To jj
    
        If j = jj Then
            Q_ARR(l + 1, j) = Q_ARR(l, j - 1) * HW_PROBABILITIES_TREE_FUNC(1, j - 1, JMAX_VAL, _
                    KAPPA, TEMP_DELTA) * Exp(-R_ARR(l, j - 1) * TEMP_DELTA)
        ElseIf j = jj - 1 Then
            Q_ARR(l + 1, j) = Q_ARR(l, j) * HW_PROBABILITIES_TREE_FUNC(0, j, JMAX_VAL, KAPPA, _
                        TEMP_DELTA) * Exp(-R_ARR(l, j) * TEMP_DELTA) + _
                        Q_ARR(l, j - 1) * HW_PROBABILITIES_TREE_FUNC(1, j - 1, JMAX_VAL, _
                        KAPPA, TEMP_DELTA) * Exp(-R_ARR(l, j - 1) * TEMP_DELTA)
        ElseIf j < jj - 1 And j > ii + 1 Then
            Q_ARR(l + 1, j) = Q_ARR(l, j - 1) * HW_PROBABILITIES_TREE_FUNC(1, j - 1, _
                        JMAX_VAL, KAPPA, TEMP_DELTA) * Exp(-R_ARR(l, j - 1) * _
                        TEMP_DELTA) + Q_ARR(l, j) * HW_PROBABILITIES_TREE_FUNC(0, j, JMAX_VAL, _
                        KAPPA, TEMP_DELTA) * Exp(-R_ARR(l, j) * TEMP_DELTA) + _
                        Q_ARR(l, j + 1) * HW_PROBABILITIES_TREE_FUNC(-1, j + 1, JMAX_VAL, KAPPA, _
                        TEMP_DELTA) * Exp(-R_ARR(l, j + 1) * TEMP_DELTA)
        ElseIf j = ii + 1 Then
            Q_ARR(l + 1, j) = Q_ARR(l, j) * HW_PROBABILITIES_TREE_FUNC(0, j, JMAX_VAL, KAPPA, _
                            TEMP_DELTA) * Exp(-R_ARR(l, j) * TEMP_DELTA) + _
                            Q_ARR(l, j + 1) * HW_PROBABILITIES_TREE_FUNC(-1, j + 1, JMAX_VAL, _
                            KAPPA, TEMP_DELTA) * Exp(-R_ARR(l, j + 1) * TEMP_DELTA)
        ElseIf j = ii Then
            Q_ARR(l + 1, j) = Q_ARR(l, j + 1) * HW_PROBABILITIES_TREE_FUNC(-1, j + 1, JMAX_VAL, _
                            KAPPA, TEMP_DELTA) * Exp(-R_ARR(l, j + 1) * TEMP_DELTA)
        End If
    Next j
Next l
    
For l = NSIZE To nSTEPS
    
    JPLUS_VAL = NSIZE
    JMINUS_VAL = -NSIZE
    
    TEMP_SUM = 0
    For j = -NSIZE To NSIZE
        TEMP_SUM = TEMP_SUM + Q_ARR(l, j) * Exp(-j * BTEMP_VAL * TEMP_DELTA)
    Next j
    
    ALPHA_VECTOR(l) = (Log(TEMP_SUM) - Log(HW_PRICE_FUNC(TENOR_RNG, _
                    ZERO_RNG, (l + 1) * TEMP_DELTA))) / TEMP_DELTA
    For j = JMINUS_VAL To JPLUS_VAL
        R_ARR(l, j) = ALPHA_VECTOR(l) + j * BTEMP_VAL
        S_MATRIX(l, j) = HW_CONTINUOUS_RATE_FUNC(R_ARR(l, j), TEMP_DELTA, l * _
                        TEMP_DELTA, SIGMA, KAPPA, TENOR_RNG, ZERO_RNG)
    Next j
    
    For j = JMINUS_VAL To JPLUS_VAL
    
    'considering the spezial case when nonstandard branching leads to central
    'nodes with five incoming arrows
        If JMAX_VAL = 2 Then
            Q_ARR(l + 1, 2) = Q_ARR(l, 2) * HW_PROBABILITIES_TREE_FUNC(1, 2, JMAX_VAL, _
                            KAPPA, TEMP_DELTA) * Exp(-R_ARR(l, 2) * TEMP_DELTA) + _
                            Q_ARR(l, 1) * HW_PROBABILITIES_TREE_FUNC(1, 1, JMAX_VAL, KAPPA, TEMP_DELTA) _
                            * Exp(-R_ARR(l, 1) * TEMP_DELTA)
            Q_ARR(l + 1, 1) = Q_ARR(l, 2) * HW_PROBABILITIES_TREE_FUNC(0, 2, JMAX_VAL, _
                            KAPPA, TEMP_DELTA) * Exp(-R_ARR(l, 2) * TEMP_DELTA) + _
                            Q_ARR(l, 1) * HW_PROBABILITIES_TREE_FUNC(0, 1, JMAX_VAL, KAPPA, TEMP_DELTA) _
                            * Exp(-R_ARR(l, 1) * TEMP_DELTA) + Q_ARR(l, 0) * _
                            HW_PROBABILITIES_TREE_FUNC(1, 0, JMAX_VAL, KAPPA, TEMP_DELTA) * Exp(-R_ARR(l, 0) _
                            * TEMP_DELTA)
            Q_ARR(l + 1, 0) = Q_ARR(l, 2) * HW_PROBABILITIES_TREE_FUNC(-1, 2, JMAX_VAL, _
                            KAPPA, TEMP_DELTA) * Exp(-R_ARR(l, 2) * TEMP_DELTA) + _
                            Q_ARR(l, 1) * HW_PROBABILITIES_TREE_FUNC(-1, 1, JMAX_VAL, KAPPA, TEMP_DELTA) _
                            * Exp(-R_ARR(l, 1) * TEMP_DELTA) + Q_ARR(l, 0) * _
                            HW_PROBABILITIES_TREE_FUNC(0, 0, JMAX_VAL, KAPPA, TEMP_DELTA) * Exp(-R_ARR(l, 0) _
                            * TEMP_DELTA) + Q_ARR(l, -1) * HW_PROBABILITIES_TREE_FUNC(1, -1, JMAX_VAL, _
                            KAPPA, TEMP_DELTA) * Exp(-R_ARR(l, -1) * TEMP_DELTA) + _
                            Q_ARR(l, -2) * HW_PROBABILITIES_TREE_FUNC(1, -2, JMAX_VAL, KAPPA, TEMP_DELTA) _
                            * Exp(-R_ARR(l, -2) * TEMP_DELTA)
            Q_ARR(l + 1, -1) = Q_ARR(l, 0) * HW_PROBABILITIES_TREE_FUNC(-1, 0, JMAX_VAL, _
                            KAPPA, TEMP_DELTA) * Exp(-R_ARR(l, 0) * TEMP_DELTA) + _
                            Q_ARR(l, -1) * HW_PROBABILITIES_TREE_FUNC(0, -1, JMAX_VAL, KAPPA, _
                            TEMP_DELTA) * Exp(-R_ARR(l, -1) * TEMP_DELTA) + _
                            Q_ARR(l, -2) * HW_PROBABILITIES_TREE_FUNC(1, -2, JMAX_VAL, KAPPA, _
                            TEMP_DELTA) * Exp(-R_ARR(l, -2) * TEMP_DELTA)
            Q_ARR(l + 1, -2) = Q_ARR(l, -1) * HW_PROBABILITIES_TREE_FUNC(-1, -1, JMAX_VAL, _
                            KAPPA, TEMP_DELTA) * Exp(-R_ARR(l, -1) * TEMP_DELTA) _
                            + Q_ARR(l, -2) * HW_PROBABILITIES_TREE_FUNC(-1, -2, JMAX_VAL, KAPPA, _
                            TEMP_DELTA) * Exp(-R_ARR(l, -2) * TEMP_DELTA)
        Else
            If j = JPLUS_VAL Then
                Q_ARR(l + 1, j) = Q_ARR(l, j) * HW_PROBABILITIES_TREE_FUNC(1, j, JMAX_VAL, _
                                KAPPA, TEMP_DELTA) * Exp(-R_ARR(l, j) * _
                                TEMP_DELTA) + Q_ARR(l, j - 1) * HW_PROBABILITIES_TREE_FUNC(1, _
                                j - 1, JMAX_VAL, KAPPA, TEMP_DELTA) * Exp(-R_ARR(l, _
                                j - 1) * TEMP_DELTA)
            ElseIf j = JPLUS_VAL - 1 Then
                Q_ARR(l + 1, j) = Q_ARR(l, j + 1) * HW_PROBABILITIES_TREE_FUNC(0, j + 1, JMAX_VAL, _
                                KAPPA, TEMP_DELTA) * Exp(-R_ARR(l, j + 1) * _
                                TEMP_DELTA) + Q_ARR(l, j) * HW_PROBABILITIES_TREE_FUNC(0, j, _
                                JMAX_VAL, KAPPA, TEMP_DELTA) * Exp(-R_ARR(l, j) * _
                                TEMP_DELTA) + Q_ARR(l, j - 1) * HW_PROBABILITIES_TREE_FUNC(1, _
                                j - 1, JMAX_VAL, KAPPA, TEMP_DELTA) * Exp(-R_ARR(l, _
                                j - 1) * TEMP_DELTA)
            ElseIf j = JPLUS_VAL - 2 Then
                Q_ARR(l + 1, j) = Q_ARR(l, j + 1) * HW_PROBABILITIES_TREE_FUNC(-1, j + 1, JMAX_VAL, _
                                KAPPA, TEMP_DELTA) * Exp(-R_ARR(l, j + 1) * TEMP_DELTA) _
                                + Q_ARR(l, j) * HW_PROBABILITIES_TREE_FUNC(0, j, JMAX_VAL, KAPPA, _
                                TEMP_DELTA) * Exp(-R_ARR(l, j) * TEMP_DELTA) + _
                                Q_ARR(l, j - 1) * HW_PROBABILITIES_TREE_FUNC(1, j - 1, JMAX_VAL, _
                                KAPPA, TEMP_DELTA) * Exp(-R_ARR(l, j - 1) * _
                                TEMP_DELTA) + Q_ARR(l, j + 2) * HW_PROBABILITIES_TREE_FUNC(-1, _
                                j + 2, JMAX_VAL, KAPPA, TEMP_DELTA) * Exp(-R_ARR(l, _
                                j + 2) * TEMP_DELTA)
            ElseIf j < JPLUS_VAL - 2 And j > JMINUS_VAL + 2 Then
                Q_ARR(l + 1, j) = Q_ARR(l, j + 1) * HW_PROBABILITIES_TREE_FUNC(-1, j + 1, JMAX_VAL, _
                                KAPPA, TEMP_DELTA) * Exp(-R_ARR(l, j + 1) * TEMP_DELTA) _
                                + Q_ARR(l, j) * HW_PROBABILITIES_TREE_FUNC(0, j, JMAX_VAL, KAPPA, TEMP_DELTA) _
                                * Exp(-R_ARR(l, j) * TEMP_DELTA) + Q_ARR(l, j - 1) * _
                                HW_PROBABILITIES_TREE_FUNC(1, j - 1, JMAX_VAL, KAPPA, TEMP_DELTA) * _
                                Exp(-R_ARR(l, j - 1) * TEMP_DELTA)
            ElseIf j = JMINUS_VAL + 2 Then
                Q_ARR(l + 1, j) = Q_ARR(l, j + 1) * HW_PROBABILITIES_TREE_FUNC(-1, j + 1, JMAX_VAL, KAPPA, _
                                TEMP_DELTA) * Exp(-R_ARR(l, j + 1) * TEMP_DELTA) + _
                                Q_ARR(l, j) * HW_PROBABILITIES_TREE_FUNC(0, j, JMAX_VAL, KAPPA, TEMP_DELTA) _
                                * Exp(-R_ARR(l, j) * TEMP_DELTA) + Q_ARR(l, j - 1) * _
                                HW_PROBABILITIES_TREE_FUNC(1, j - 1, JMAX_VAL, KAPPA, TEMP_DELTA) * _
                                Exp(-R_ARR(l, j - 1) * TEMP_DELTA) + Q_ARR(l, j - 2) * _
                                HW_PROBABILITIES_TREE_FUNC(1, j - 2, JMAX_VAL, KAPPA, TEMP_DELTA) * _
                                Exp(-R_ARR(l, j - 2) * TEMP_DELTA)
            ElseIf j = JMINUS_VAL + 1 Then
                Q_ARR(l + 1, j) = Q_ARR(l, j) * HW_PROBABILITIES_TREE_FUNC(0, j, JMAX_VAL, KAPPA, TEMP_DELTA) _
                                * Exp(-R_ARR(l, j) * TEMP_DELTA) + Q_ARR(l, j - 1) * _
                                HW_PROBABILITIES_TREE_FUNC(0, j - 1, JMAX_VAL, KAPPA, TEMP_DELTA) * _
                                Exp(-R_ARR(l, j - 1) * TEMP_DELTA) + Q_ARR(l, j + 1) * _
                                HW_PROBABILITIES_TREE_FUNC(-1, j + 1, JMAX_VAL, KAPPA, TEMP_DELTA) * _
                                Exp(-R_ARR(l, j + 1) * TEMP_DELTA)
            ElseIf j = JMINUS_VAL Then
                Q_ARR(l + 1, j) = Q_ARR(l, j) * HW_PROBABILITIES_TREE_FUNC(-1, j, JMAX_VAL, KAPPA, TEMP_DELTA) _
                                * Exp(-R_ARR(l, j) * TEMP_DELTA) + Q_ARR(l, j + 1) * _
                                HW_PROBABILITIES_TREE_FUNC(-1, j + 1, JMAX_VAL, KAPPA, TEMP_DELTA) * _
                                Exp(-R_ARR(l, j + 1) * TEMP_DELTA)
            End If
        End If
    Next j
    
Next l

'cash flow calculations
For k = 1 To nSTEPS
    If k / TAU_VAL - Int(k / TAU_VAL) <> 0 Then
        S_VECTOR(k) = (Int(k / TAU_VAL) + 1) * TAU_VAL 'time of next swap payment date in days
    Else: S_VECTOR(k) = 0
    End If
    For j = -MINIMUM_FUNC(k, JMAX_VAL) To MINIMUM_FUNC(k, JMAX_VAL)
        If HW_ZERO_FUNC(HW_CONTINUOUS_RATE_FUNC(R_ARR(k, j), TEMP_DELTA, k * _
                TEMP_DELTA, SIGMA, KAPPA, TENOR_RNG, ZERO_RNG), k * _
                TEMP_DELTA, k * TEMP_DELTA + TAU_VAL / TRADING_DAYS, SIGMA, _
                KAPPA, TENOR_RNG, ZERO_RNG) < 0.08 Then
            
            O_ARR(k, j) = (CASH_FLOW / TRADING_DAYS * TAU_VAL / TRADING_DAYS * _
                        (Exp(STRIKE_RATE * TAU_VAL / TRADING_DAYS) - 1)) * _
                        Exp(-HW_ZERO_FUNC(HW_CONTINUOUS_RATE_FUNC(R_ARR(k, j), TEMP_DELTA, _
                        k * TEMP_DELTA, SIGMA, KAPPA, TENOR_RNG, ZERO_RNG), k * _
                        TEMP_DELTA, S_VECTOR(k) / TRADING_DAYS, SIGMA, KAPPA, _
                        TENOR_RNG, ZERO_RNG) * (S_VECTOR(k) - k)) / TRADING_DAYS _
                        + O_ARR(k, j)
        Else: O_ARR(k, j) = O_ARR(k, j)
        End If
    Next j
Next k

'backward induction
For i = nSTEPS To NSIZE + 1 Step -1
    For j = -NSIZE To NSIZE
        If j = JMAX_VAL Then
            O_ARR(i - 1, j) = (O_ARR(i, j) * HW_PROBABILITIES_TREE_FUNC(1, j, JMAX_VAL, KAPPA, _
                            TEMP_DELTA) + O_ARR(i, j - 1) * HW_PROBABILITIES_TREE_FUNC(0, j, _
                            JMAX_VAL, KAPPA, TEMP_DELTA) + O_ARR(i, j - 2) * _
                            HW_PROBABILITIES_TREE_FUNC(-1, j, JMAX_VAL, KAPPA, TEMP_DELTA)) * _
                            Exp(-R_ARR(i - 1, j) * TEMP_DELTA) + O_ARR(i - 1, j)
        ElseIf j < JMAX_VAL And j > JMIN_VAL Then
            O_ARR(i - 1, j) = (O_ARR(i, j + 1) * HW_PROBABILITIES_TREE_FUNC(1, j, JMAX_VAL, _
                            KAPPA, TEMP_DELTA) + O_ARR(i, j) * HW_PROBABILITIES_TREE_FUNC(0, j, _
                            JMAX_VAL, KAPPA, TEMP_DELTA) + O_ARR(i, j - 1) * _
                            HW_PROBABILITIES_TREE_FUNC(-1, j, JMAX_VAL, KAPPA, TEMP_DELTA)) * _
                            Exp(-R_ARR(i - 1, j) * TEMP_DELTA) + O_ARR(i - 1, j)
        ElseIf j = JMIN_VAL Then
            O_ARR(i - 1, j) = (O_ARR(i, j) * HW_PROBABILITIES_TREE_FUNC(-1, j, JMAX_VAL, _
                            KAPPA, TEMP_DELTA) + O_ARR(i, j + 1) * HW_PROBABILITIES_TREE_FUNC(0, _
                            j, JMAX_VAL, KAPPA, TEMP_DELTA) + O_ARR(i, j + 2) * _
                            HW_PROBABILITIES_TREE_FUNC(1, j, JMAX_VAL, KAPPA, TEMP_DELTA)) * _
                            Exp(-R_ARR(i - 1, j) * TEMP_DELTA) + O_ARR(i - 1, j)
        End If
    Next j
Next i

JMINUS_VAL = -NSIZE
JPLUS_VAL = NSIZE

For i = NSIZE To 1 Step -1
    JMINUS_VAL = JMINUS_VAL + 1
    JPLUS_VAL = JPLUS_VAL - 1
    For j = JMINUS_VAL To JPLUS_VAL
        O_ARR(i - 1, j) = (O_ARR(i, j + 1) * HW_PROBABILITIES_TREE_FUNC(1, j, JMAX_VAL, _
                        KAPPA, TEMP_DELTA) + O_ARR(i, j) * HW_PROBABILITIES_TREE_FUNC(0, j, _
                        JMAX_VAL, KAPPA, TEMP_DELTA) + O_ARR(i, j - 1) * HW_PROBABILITIES_TREE_FUNC(-1, _
                        j, JMAX_VAL, KAPPA, TEMP_DELTA)) * Exp(-R_ARR(i - 1, j) * _
                        TEMP_DELTA) + O_ARR(i - 1, j)
    Next j
Next i


HW_BINARY_FUNC = O_ARR(0, 0)
   
Exit Function
ERROR_LABEL:
    HW_BINARY_FUNC = Err.number
End Function

'**********************************************************************************
'**********************************************************************************
'FUNCTION      : HW_AMERICAN_OPTION_FUNC

'DESCRIPTION   : Valuation of American style option
'If we assume the discount bond option to be of American style we
'calculate the terminal payoff of the option in the usual way, but
'in the backward induction process we have to consider that early
'exercise may be optimal. This means that we work backward in the tree.

'However it is straightforward that it will never be optimal to exercise
'an American call option as the value of the bond is bound to increase and
'it will always be optimal to exercise an American put option immediately.

'LIBRARY       : FIXED_INCOME
'GROUP         : HW_1F
'ID            : 012
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'REFERENCE: http://www.angelfire.com/ny/financeinfo/HullWhite.pdf
'**********************************************************************************
'**********************************************************************************

Function HW_AMERICAN_OPTION_FUNC(ByVal OPTION_FLAG As Integer, _
ByVal DELTA As Double, _
ByVal SIGMA As Double, _
ByVal KAPPA As Double, _
ByRef FIRST_TENOR_RNG As Variant, _
ByVal SECOND_TENOR As Double, _
ByVal CASH_FLOW As Double, _
ByVal STRIKE As Double, _
ByRef TENOR_RNG As Variant, _
ByRef ZERO_RNG As Variant)

Dim h As Double
Dim i As Double
Dim j As Double
Dim k As Double
Dim l As Double

Dim ii As Double
Dim jj As Double

Dim JMAX_VAL As Double
Dim JMIN_VAL As Double

Dim JPLUS_VAL As Double
Dim JMINUS_VAL As Double

Dim NSIZE As Double
Dim NROWS As Double
Dim nSTEPS As Double

Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double
Dim CTEMP_VAL As Double
Dim DTEMP_VAL As Double

Dim TEMP_SUM As Double

Dim P_ARR As Variant
Dim Q_ARR As Variant
Dim O_ARR As Variant
Dim R_ARR As Variant
Dim S_ARR As Variant
Dim T_ARR As Variant

Dim ALPHA_VECTOR As Variant
Dim STEPS_VECTOR As Variant
Dim FIRST_TENOR_VECTOR As Variant

On Error GoTo ERROR_LABEL

If IsArray(FIRST_TENOR_RNG) = True Then
    FIRST_TENOR_VECTOR = FIRST_TENOR_RNG
    If UBound(FIRST_TENOR_VECTOR, 1) = 1 Then: _
        FIRST_TENOR_VECTOR = MATRIX_TRANSPOSE_FUNC(FIRST_TENOR_VECTOR)
Else
    ReDim FIRST_TENOR_VECTOR(1 To 1, 1 To 1)
    FIRST_TENOR_VECTOR(1, 1) = FIRST_TENOR_RNG
End If


Select Case OPTION_FLAG
Case -1 '"put", "p", -1
      h = 1
Case Else
      h = -1
End Select

NROWS = UBound(FIRST_TENOR_VECTOR, 1)

If FIRST_TENOR_VECTOR(NROWS, 1) / DELTA - _
    Int(FIRST_TENOR_VECTOR(NROWS, 1) / DELTA) = 0 Then
    nSTEPS = FIRST_TENOR_VECTOR(NROWS, 1) / DELTA
Else
    nSTEPS = Int(FIRST_TENOR_VECTOR(NROWS, 1) / DELTA) + 1
End If

'calculating constants
ATEMP_VAL = SIGMA ^ 2 * (1 - Exp(-2 * KAPPA * DELTA)) / (2 * KAPPA)
BTEMP_VAL = Sqr(3 * ATEMP_VAL)
JMAX_VAL = Int(0.184 / (KAPPA * DELTA)) + 1
JMIN_VAL = -JMAX_VAL
NSIZE = MINIMUM_FUNC(nSTEPS, JMAX_VAL)

'redimensioning variables
ReDim ALPHA_VECTOR(0 To nSTEPS)
ReDim Q_ARR(0 To nSTEPS + 1, -NSIZE To NSIZE)
ReDim O_ARR(0 To nSTEPS, -NSIZE To NSIZE)
ReDim R_ARR(0 To nSTEPS, -NSIZE To NSIZE)
ReDim S_ARR(0 To nSTEPS, -NSIZE To NSIZE)
ReDim P_ARR(0 To nSTEPS, -NSIZE To NSIZE)
ReDim T_ARR(0 To nSTEPS, -NSIZE To NSIZE)
ReDim STEPS_VECTOR(0 To NROWS)

'calculating start values
ALPHA_VECTOR(0) = -Log(HW_PRICE_FUNC(TENOR_RNG, ZERO_RNG, DELTA)) / DELTA
R_ARR(0, 0) = ALPHA_VECTOR(0)
S_ARR(0, 0) = HW_CONTINUOUS_RATE_FUNC(R_ARR(0, 0), DELTA, 0, SIGMA, _
                KAPPA, TENOR_RNG, ZERO_RNG)
P_ARR(0, 0) = HW_A_FUNC(0 * DELTA, SECOND_TENOR, SIGMA, KAPPA, _
                TENOR_RNG, ZERO_RNG) * Exp(-HW_B_FUNC(0 * DELTA, _
                SECOND_TENOR, KAPPA) * S_ARR(0, 0)) * CASH_FLOW
T_ARR(0, 0) = MAXIMUM_FUNC(h * (STRIKE - P_ARR(0, 0)), 0)
Q_ARR(1, 1) = HW_PROBABILITIES_TREE_FUNC(1, 0, JMAX_VAL, KAPPA, DELTA) * Exp(-R_ARR(0, 0) * DELTA)
Q_ARR(1, 0) = HW_PROBABILITIES_TREE_FUNC(0, 0, JMAX_VAL, KAPPA, DELTA) * Exp(-R_ARR(0, 0) * DELTA)
Q_ARR(1, -1) = HW_PROBABILITIES_TREE_FUNC(-1, 0, JMAX_VAL, KAPPA, DELTA) * Exp(-R_ARR(0, 0) * DELTA)

JMINUS_VAL = 0
JPLUS_VAL = 0

ii = -1
jj = 1

'using Q_ARR(l) to calculate ALPHA_VECTOR(l),R_ARR(l) and then
'Q_ARR(m+1) up to NSIZE-1 so we get the Q_ARR(NSIZE) as last results

For l = 1 To NSIZE - 1
    
    JMINUS_VAL = JMINUS_VAL - 1
    JPLUS_VAL = JPLUS_VAL + 1
    
    TEMP_SUM = 0
    For j = JMINUS_VAL To JPLUS_VAL
        TEMP_SUM = TEMP_SUM + Q_ARR(l, j) * Exp(-j * BTEMP_VAL * DELTA)
    Next j
    
    ALPHA_VECTOR(l) = (Log(TEMP_SUM) - Log(HW_PRICE_FUNC(TENOR_RNG, _
                    ZERO_RNG, (l + 1) * DELTA))) / DELTA
    For j = JMINUS_VAL To JPLUS_VAL
        R_ARR(l, j) = ALPHA_VECTOR(l) + j * BTEMP_VAL
        S_ARR(l, j) = HW_CONTINUOUS_RATE_FUNC(R_ARR(l, j), DELTA, l * DELTA, _
                    SIGMA, KAPPA, TENOR_RNG, ZERO_RNG)
        P_ARR(l, j) = HW_A_FUNC(l * DELTA, SECOND_TENOR, SIGMA, KAPPA, _
                        TENOR_RNG, ZERO_RNG) * Exp(-HW_B_FUNC(l * DELTA, _
                        SECOND_TENOR, KAPPA) * S_ARR(l, j)) * CASH_FLOW
        T_ARR(l, j) = MAXIMUM_FUNC(h * (STRIKE - P_ARR(l, j)), 0)
    Next j
    
    ii = ii - 1
    jj = jj + 1
    
    For j = ii To jj
        If j = jj Then
            Q_ARR(l + 1, j) = Q_ARR(l, j - 1) * HW_PROBABILITIES_TREE_FUNC(1, j - 1, JMAX_VAL, _
                    KAPPA, DELTA) * Exp(-R_ARR(l, j - 1) * DELTA)
        ElseIf j = jj - 1 Then
            Q_ARR(l + 1, j) = Q_ARR(l, j) * HW_PROBABILITIES_TREE_FUNC(0, j, JMAX_VAL, KAPPA, _
                    DELTA) * Exp(-R_ARR(l, j) * DELTA) + Q_ARR(l, j - 1) * _
                    HW_PROBABILITIES_TREE_FUNC(1, j - 1, JMAX_VAL, KAPPA, DELTA) * Exp(-R_ARR(l, _
                    j - 1) * DELTA)
        ElseIf j < jj - 1 And j > ii + 1 Then
            Q_ARR(l + 1, j) = Q_ARR(l, j - 1) * HW_PROBABILITIES_TREE_FUNC(1, j - 1, JMAX_VAL, _
                    KAPPA, DELTA) * Exp(-R_ARR(l, j - 1) * DELTA) + _
                    Q_ARR(l, j) * HW_PROBABILITIES_TREE_FUNC(0, j, JMAX_VAL, KAPPA, DELTA) * _
                    Exp(-R_ARR(l, j) * DELTA) + Q_ARR(l, j + 1) * HW_PROBABILITIES_TREE_FUNC(-1, _
                    j + 1, JMAX_VAL, KAPPA, DELTA) * Exp(-R_ARR(l, j + 1) * DELTA)
        ElseIf j = ii + 1 Then
            Q_ARR(l + 1, j) = Q_ARR(l, j) * HW_PROBABILITIES_TREE_FUNC(0, j, JMAX_VAL, KAPPA, DELTA) * _
                    Exp(-R_ARR(l, j) * DELTA) + Q_ARR(l, j + 1) * HW_PROBABILITIES_TREE_FUNC(-1, _
                    j + 1, JMAX_VAL, KAPPA, DELTA) * Exp(-R_ARR(l, j + 1) * DELTA)
        ElseIf j = ii Then
            Q_ARR(l + 1, j) = Q_ARR(l, j + 1) * HW_PROBABILITIES_TREE_FUNC(-1, j + 1, JMAX_VAL, KAPPA, _
                            DELTA) * Exp(-R_ARR(l, j + 1) * DELTA)
        End If
    Next j
Next l
    
    'using the Q_ARR(NSIZE) to calculate ALPHA_VECTOR(NSIZE) and so on up to NSTEPS
For l = NSIZE To nSTEPS
    
    JPLUS_VAL = NSIZE
    JMINUS_VAL = -NSIZE
    
    TEMP_SUM = 0
    For j = -NSIZE To NSIZE
        TEMP_SUM = TEMP_SUM + Q_ARR(l, j) * Exp(-j * BTEMP_VAL * DELTA)
    Next j
    
    ALPHA_VECTOR(l) = (Log(TEMP_SUM) - Log(HW_PRICE_FUNC(TENOR_RNG, ZERO_RNG, _
                    (l + 1) * DELTA))) / DELTA
    For j = JMINUS_VAL To JPLUS_VAL
        R_ARR(l, j) = ALPHA_VECTOR(l) + j * BTEMP_VAL
        S_ARR(l, j) = HW_CONTINUOUS_RATE_FUNC(R_ARR(l, j), DELTA, l * DELTA, SIGMA, _
                        KAPPA, TENOR_RNG, ZERO_RNG)
        P_ARR(l, j) = HW_A_FUNC(l * DELTA, SECOND_TENOR, SIGMA, KAPPA, _
                    TENOR_RNG, ZERO_RNG) * Exp(-HW_B_FUNC(l * DELTA, _
                    SECOND_TENOR, KAPPA) * S_ARR(l, j)) * CASH_FLOW
        T_ARR(l, j) = MAXIMUM_FUNC(h * (STRIKE - P_ARR(l, j)), 0)
    Next j
    
    For j = JMINUS_VAL To JPLUS_VAL
    
    'considering the spezial case when nonstandard branching leads to central _
     nodes with five incoming arrows
        If JMAX_VAL = 2 Then
            Q_ARR(l + 1, 2) = Q_ARR(l, 2) * HW_PROBABILITIES_TREE_FUNC(1, 2, JMAX_VAL, _
                        KAPPA, DELTA) * Exp(-R_ARR(l, 2) * DELTA) + _
                        Q_ARR(l, 1) * HW_PROBABILITIES_TREE_FUNC(1, 1, JMAX_VAL, KAPPA, _
                        DELTA) * Exp(-R_ARR(l, 1) * DELTA)
            Q_ARR(l + 1, 1) = Q_ARR(l, 2) * HW_PROBABILITIES_TREE_FUNC(0, 2, JMAX_VAL, KAPPA, _
                        DELTA) * Exp(-R_ARR(l, 2) * DELTA) + Q_ARR(l, 1) * _
                        HW_PROBABILITIES_TREE_FUNC(0, 1, JMAX_VAL, KAPPA, DELTA) * Exp(-R_ARR(l, 1) * _
                        DELTA) + Q_ARR(l, 0) * HW_PROBABILITIES_TREE_FUNC(1, 0, JMAX_VAL, KAPPA, _
                        DELTA) * Exp(-R_ARR(l, 0) * DELTA)
            Q_ARR(l + 1, 0) = Q_ARR(l, 2) * HW_PROBABILITIES_TREE_FUNC(-1, 2, JMAX_VAL, KAPPA, DELTA) * _
                        Exp(-R_ARR(l, 2) * DELTA) + Q_ARR(l, 1) * HW_PROBABILITIES_TREE_FUNC(-1, _
                        1, JMAX_VAL, KAPPA, DELTA) * Exp(-R_ARR(l, 1) * DELTA) + _
                        Q_ARR(l, 0) * HW_PROBABILITIES_TREE_FUNC(0, 0, JMAX_VAL, KAPPA, DELTA) * _
                        Exp(-R_ARR(l, 0) * DELTA) + Q_ARR(l, -1) * HW_PROBABILITIES_TREE_FUNC(1, _
                        -1, JMAX_VAL, KAPPA, DELTA) * Exp(-R_ARR(l, -1) * DELTA) + _
                        Q_ARR(l, -2) * HW_PROBABILITIES_TREE_FUNC(1, -2, JMAX_VAL, KAPPA, DELTA) * _
                        Exp(-R_ARR(l, -2) * DELTA)
            Q_ARR(l + 1, -1) = Q_ARR(l, 0) * HW_PROBABILITIES_TREE_FUNC(-1, 0, JMAX_VAL, KAPPA, DELTA) * _
                        Exp(-R_ARR(l, 0) * DELTA) + Q_ARR(l, -1) * HW_PROBABILITIES_TREE_FUNC(0, _
                        -1, JMAX_VAL, KAPPA, DELTA) * Exp(-R_ARR(l, -1) * DELTA) + _
                        Q_ARR(l, -2) * HW_PROBABILITIES_TREE_FUNC(1, -2, JMAX_VAL, KAPPA, DELTA) * _
                        Exp(-R_ARR(l, -2) * DELTA)
            Q_ARR(l + 1, -2) = Q_ARR(l, -1) * HW_PROBABILITIES_TREE_FUNC(-1, -1, JMAX_VAL, KAPPA, DELTA) * _
                        Exp(-R_ARR(l, -1) * DELTA) + Q_ARR(l, -2) * HW_PROBABILITIES_TREE_FUNC(-1, _
                        -2, JMAX_VAL, KAPPA, DELTA) * Exp(-R_ARR(l, -2) * DELTA)
        Else
            If j = JPLUS_VAL Then
                Q_ARR(l + 1, j) = Q_ARR(l, j) * HW_PROBABILITIES_TREE_FUNC(1, j, JMAX_VAL, _
                                KAPPA, DELTA) * Exp(-R_ARR(l, j) * DELTA) + _
                                Q_ARR(l, j - 1) * HW_PROBABILITIES_TREE_FUNC(1, j - 1, JMAX_VAL, _
                                KAPPA, DELTA) * Exp(-R_ARR(l, j - 1) * DELTA)
            ElseIf j = JPLUS_VAL - 1 Then
                Q_ARR(l + 1, j) = Q_ARR(l, j + 1) * HW_PROBABILITIES_TREE_FUNC(0, j + 1, _
                                JMAX_VAL, KAPPA, DELTA) * Exp(-R_ARR(l, j + 1) * _
                                DELTA) + Q_ARR(l, j) * HW_PROBABILITIES_TREE_FUNC(0, j, JMAX_VAL, _
                                KAPPA, DELTA) * Exp(-R_ARR(l, j) * DELTA) + _
                                Q_ARR(l, j - 1) * HW_PROBABILITIES_TREE_FUNC(1, j - 1, JMAX_VAL, _
                                KAPPA, DELTA) * Exp(-R_ARR(l, j - 1) * DELTA)
            ElseIf j = JPLUS_VAL - 2 Then
                Q_ARR(l + 1, j) = Q_ARR(l, j + 1) * HW_PROBABILITIES_TREE_FUNC(-1, j + 1, JMAX_VAL, _
                                KAPPA, DELTA) * Exp(-R_ARR(l, j + 1) * DELTA) + _
                                Q_ARR(l, j) * HW_PROBABILITIES_TREE_FUNC(0, j, JMAX_VAL, KAPPA, _
                                DELTA) * Exp(-R_ARR(l, j) * DELTA) + Q_ARR(l, j - 1) * _
                                HW_PROBABILITIES_TREE_FUNC(1, j - 1, JMAX_VAL, KAPPA, DELTA) * _
                                Exp(-R_ARR(l, j - 1) * DELTA) + Q_ARR(l, j + 2) * _
                                HW_PROBABILITIES_TREE_FUNC(-1, j + 2, JMAX_VAL, KAPPA, DELTA) * _
                                Exp(-R_ARR(l, j + 2) * DELTA)
            ElseIf j < JPLUS_VAL - 2 And j > JMINUS_VAL + 2 Then
                Q_ARR(l + 1, j) = Q_ARR(l, j + 1) * HW_PROBABILITIES_TREE_FUNC(-1, j + 1, JMAX_VAL, _
                                KAPPA, DELTA) * Exp(-R_ARR(l, j + 1) * DELTA) + _
                                Q_ARR(l, j) * HW_PROBABILITIES_TREE_FUNC(0, j, JMAX_VAL, KAPPA, _
                                DELTA) * Exp(-R_ARR(l, j) * DELTA) + Q_ARR(l, j - 1) _
                                * HW_PROBABILITIES_TREE_FUNC(1, j - 1, JMAX_VAL, KAPPA, DELTA) * _
                                Exp(-R_ARR(l, j - 1) * DELTA)
            ElseIf j = JMINUS_VAL + 2 Then
                Q_ARR(l + 1, j) = Q_ARR(l, j + 1) * HW_PROBABILITIES_TREE_FUNC(-1, j + 1, JMAX_VAL, _
                                KAPPA, DELTA) * Exp(-R_ARR(l, j + 1) * DELTA) + _
                                Q_ARR(l, j) * HW_PROBABILITIES_TREE_FUNC(0, j, JMAX_VAL, KAPPA, _
                                DELTA) * Exp(-R_ARR(l, j) * DELTA) + Q_ARR(l, j - 1) _
                                * HW_PROBABILITIES_TREE_FUNC(1, j - 1, JMAX_VAL, KAPPA, DELTA) * _
                                Exp(-R_ARR(l, j - 1) * DELTA) + Q_ARR(l, j - 2) _
                                * HW_PROBABILITIES_TREE_FUNC(1, j - 2, JMAX_VAL, KAPPA, DELTA) * _
                                Exp(-R_ARR(l, j - 2) * DELTA)
            ElseIf j = JMINUS_VAL + 1 Then
                Q_ARR(l + 1, j) = Q_ARR(l, j) * HW_PROBABILITIES_TREE_FUNC(0, j, JMAX_VAL, KAPPA, _
                                DELTA) * Exp(-R_ARR(l, j) * DELTA) + Q_ARR(l, j - 1) _
                                * HW_PROBABILITIES_TREE_FUNC(0, j - 1, JMAX_VAL, KAPPA, DELTA) * _
                                Exp(-R_ARR(l, j - 1) * DELTA) + Q_ARR(l, j + 1) * _
                                HW_PROBABILITIES_TREE_FUNC(-1, j + 1, JMAX_VAL, KAPPA, DELTA) * _
                                Exp(-R_ARR(l, j + 1) * DELTA)
            ElseIf j = JMINUS_VAL Then
                Q_ARR(l + 1, j) = Q_ARR(l, j) * HW_PROBABILITIES_TREE_FUNC(-1, j, JMAX_VAL, KAPPA, _
                                DELTA) * Exp(-R_ARR(l, j) * DELTA) + Q_ARR(l, j + 1) _
                                * HW_PROBABILITIES_TREE_FUNC(-1, j + 1, JMAX_VAL, KAPPA, DELTA) * _
                                Exp(-R_ARR(l, j + 1) * DELTA)
            End If
        End If
    Next j
    
Next l

'cash flow calculations
For k = 1 To NROWS
    If k = 1 Then
          CTEMP_VAL = FIRST_TENOR_VECTOR(k, 1)
    Else: CTEMP_VAL = FIRST_TENOR_VECTOR(k, 1) - FIRST_TENOR_VECTOR(k - 1, 1)
    End If

    DTEMP_VAL = FIRST_TENOR_VECTOR(k, 1) / DELTA - Int(FIRST_TENOR_VECTOR(k, 1) / DELTA)

    If DTEMP_VAL = 0 Then
        STEPS_VECTOR(k) = FIRST_TENOR_VECTOR(k, 1) / DELTA
        For j = -MINIMUM_FUNC(STEPS_VECTOR(k), JMAX_VAL) To MINIMUM_FUNC(STEPS_VECTOR(k), JMAX_VAL)
            O_ARR(STEPS_VECTOR(k), j) = T_ARR(STEPS_VECTOR(k), j) + O_ARR(STEPS_VECTOR(k), j)
        Next j
    Else
        STEPS_VECTOR(k) = Int(FIRST_TENOR_VECTOR(k, 1) / DELTA)
        For j = -MINIMUM_FUNC(STEPS_VECTOR(k), JMAX_VAL) To MINIMUM_FUNC(STEPS_VECTOR(k), JMAX_VAL)
            O_ARR(STEPS_VECTOR(k), j) = DTEMP_VAL * T_ARR(STEPS_VECTOR(k), j) _
                                    + O_ARR(STEPS_VECTOR(k), j)
        Next j
        For j = -MINIMUM_FUNC(STEPS_VECTOR(k) + 1, JMAX_VAL) To _
                MINIMUM_FUNC(STEPS_VECTOR(k) + 1, JMAX_VAL)
            O_ARR(STEPS_VECTOR(k) + 1, j) = (1 - DTEMP_VAL) * _
                                T_ARR(STEPS_VECTOR(k) + 1, j) + _
                                O_ARR(STEPS_VECTOR(k) + 1, j)
        Next j
    End If
Next k

'backward induction
For i = nSTEPS To NSIZE + 1 Step -1
    
    For j = -NSIZE To NSIZE
        If j = JMAX_VAL Then
            O_ARR(i - 1, j) = MAXIMUM_FUNC(T_ARR(i - 1, j), _
                        (O_ARR(i, j) * HW_PROBABILITIES_TREE_FUNC(1, j, JMAX_VAL, KAPPA, _
                        DELTA) + O_ARR(i, j - 1) * HW_PROBABILITIES_TREE_FUNC(0, j, _
                        JMAX_VAL, KAPPA, DELTA) + O_ARR(i, j - 2) * HW_PROBABILITIES_TREE_FUNC(-1, _
                        j, JMAX_VAL, KAPPA, DELTA)) * Exp(-R_ARR(i - 1, j) * DELTA)) _
                        + O_ARR(i - 1, j)
        ElseIf j < JMAX_VAL And j > JMIN_VAL Then
            O_ARR(i - 1, j) = MAXIMUM_FUNC(T_ARR(i - 1, j), (O_ARR(i, j + 1) _
                            * HW_PROBABILITIES_TREE_FUNC(1, j, JMAX_VAL, KAPPA, DELTA) + O_ARR(i, j) _
                            * HW_PROBABILITIES_TREE_FUNC(0, j, JMAX_VAL, KAPPA, DELTA) + O_ARR(i, j - 1) _
                            * HW_PROBABILITIES_TREE_FUNC(-1, j, JMAX_VAL, KAPPA, DELTA)) * _
                            Exp(-R_ARR(i - 1, j) * DELTA)) + O_ARR(i - 1, j)
        ElseIf j = JMIN_VAL Then
            O_ARR(i - 1, j) = MAXIMUM_FUNC(T_ARR(i - 1, j), (O_ARR(i, j) * _
                            HW_PROBABILITIES_TREE_FUNC(-1, j, JMAX_VAL, KAPPA, DELTA) + O_ARR(i, _
                            j + 1) * HW_PROBABILITIES_TREE_FUNC(0, j, JMAX_VAL, KAPPA, DELTA) + _
                            O_ARR(i, j + 2) * HW_PROBABILITIES_TREE_FUNC(1, j, JMAX_VAL, KAPPA, _
                            DELTA)) * Exp(-R_ARR(i - 1, j) * DELTA)) + O_ARR(i - 1, j)
        End If
    Next j
Next i

JMINUS_VAL = -NSIZE
JPLUS_VAL = NSIZE

For i = NSIZE To 1 Step -1
    JMINUS_VAL = JMINUS_VAL + 1
    JPLUS_VAL = JPLUS_VAL - 1
     For j = JMINUS_VAL To JPLUS_VAL
         O_ARR(i - 1, j) = MAXIMUM_FUNC(T_ARR(i - 1, j), _
                        (O_ARR(i, j + 1) * HW_PROBABILITIES_TREE_FUNC(1, j, JMAX_VAL, KAPPA, _
                        DELTA) + O_ARR(i, j) * HW_PROBABILITIES_TREE_FUNC(0, j, JMAX_VAL, _
                        KAPPA, DELTA) + O_ARR(i, j - 1) * HW_PROBABILITIES_TREE_FUNC(-1, j, _
                        JMAX_VAL, KAPPA, DELTA)) * Exp(-R_ARR(i - 1, j) * DELTA)) + _
                        O_ARR(i - 1, j)
     Next j
Next i


HW_AMERICAN_OPTION_FUNC = O_ARR(0, 0)
   
Exit Function
ERROR_LABEL:
    HW_AMERICAN_OPTION_FUNC = Err.number
End Function

'**********************************************************************************
'**********************************************************************************
'FUNCTION      : HW_CALLABLE_FUNC

'DESCRIPTION   : Callable Bond Valuation
'Some bonds have embedded options which give either the holder or the
'issuer certain rights. a callable bond for example gives the issuer
'the right to call the bonds at certain times. With the following
'numerical procedure it is easy to price these embedded options.
'As a simple example lets consider a discount bond with 9 years to maturity
'which gives the issuer the right to call the bonds at any time for X=75.
'This callable bond is priced by calculating the price of the bond at the
'terminal nodes which is 100 and working backward by which is similar to the
'American style option valuation. Again this formula has to be adjusted
'in the right way if non-standard branching is used.

'LIBRARY       : FIXED_INCOME
'GROUP         : HW_1F
'ID            : 013
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'REFERENCE: http://www.angelfire.com/ny/financeinfo/HullWhite.pdf
'**********************************************************************************
'**********************************************************************************

Function HW_CALLABLE_FUNC(ByVal OPTION_FLAG As Integer, _
ByVal TEMP_DELTA As Double, _
ByVal SIGMA As Double, _
ByVal KAPPA As Double, _
ByVal FIRST_TENOR As Double, _
ByVal CASH_FLOW As Double, _
ByVal STRIKE As Double, _
ByVal TENOR_RNG As Variant, _
ByVal ZERO_RNG As Variant)

Dim i As Double
Dim j As Double
Dim l As Double

Dim ii As Double
Dim jj As Double

Dim JMAX_VAL As Double
Dim JMIN_VAL As Double

Dim JPLUS_VAL As Double
Dim JMINUS_VAL As Double

Dim NSIZE As Double
Dim nSTEPS As Double

Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double

Dim TEMP_SUM As Double

Dim P_ARR As Variant
Dim Q_ARR As Variant
Dim O_ARR As Variant
Dim R_ARR As Variant
Dim S_ARR As Variant

Dim ALPHA_VECTOR As Variant

On Error GoTo ERROR_LABEL

nSTEPS = FIRST_TENOR / TEMP_DELTA

ATEMP_VAL = SIGMA ^ 2 * (1 - Exp(-2 * KAPPA * _
        TEMP_DELTA)) / (2 * KAPPA) 'calculating constants
BTEMP_VAL = Sqr(3 * ATEMP_VAL)
JMAX_VAL = Int(0.184 / (KAPPA * TEMP_DELTA)) + 1
JMIN_VAL = -JMAX_VAL
NSIZE = MINIMUM_FUNC(nSTEPS, JMAX_VAL)

'redimensioning variables
ReDim ALPHA_VECTOR(0 To nSTEPS)
ReDim O_ARR(0 To nSTEPS, -NSIZE To NSIZE)
ReDim P_ARR(0 To nSTEPS, -NSIZE To NSIZE)
ReDim Q_ARR(0 To nSTEPS + 1, -NSIZE To NSIZE)
ReDim R_ARR(0 To nSTEPS, -NSIZE To NSIZE)
ReDim S_ARR(0 To nSTEPS, -NSIZE To NSIZE)

ALPHA_VECTOR(0) = -Log(HW_PRICE_FUNC(TENOR_RNG, ZERO_RNG, _
                TEMP_DELTA)) / TEMP_DELTA 'calculating start values
R_ARR(0, 0) = ALPHA_VECTOR(0)
S_ARR(0, 0) = HW_CONTINUOUS_RATE_FUNC(R_ARR(0, 0), TEMP_DELTA, 0, SIGMA, _
            KAPPA, TENOR_RNG, ZERO_RNG)
P_ARR(0, 0) = HW_A_FUNC(0 * TEMP_DELTA, FIRST_TENOR, SIGMA, _
            KAPPA, TENOR_RNG, ZERO_RNG) * Exp(-HW_B_FUNC(0 * _
            TEMP_DELTA, FIRST_TENOR, KAPPA) * S_ARR(0, 0)) * CASH_FLOW
Q_ARR(1, 1) = HW_PROBABILITIES_TREE_FUNC(1, 0, JMAX_VAL, KAPPA, TEMP_DELTA) * _
            Exp(-R_ARR(0, 0) * TEMP_DELTA)
Q_ARR(1, 0) = HW_PROBABILITIES_TREE_FUNC(0, 0, JMAX_VAL, KAPPA, TEMP_DELTA) * _
            Exp(-R_ARR(0, 0) * TEMP_DELTA)
Q_ARR(1, -1) = HW_PROBABILITIES_TREE_FUNC(-1, 0, JMAX_VAL, KAPPA, TEMP_DELTA) * _
            Exp(-R_ARR(0, 0) * TEMP_DELTA)

JMINUS_VAL = 0
JPLUS_VAL = 0

ii = -1
jj = 1

For l = 1 To NSIZE - 1
    JMINUS_VAL = JMINUS_VAL - 1
    JPLUS_VAL = JPLUS_VAL + 1
    TEMP_SUM = 0
    For j = JMINUS_VAL To JPLUS_VAL
        TEMP_SUM = TEMP_SUM + Q_ARR(l, j) * Exp(-j * BTEMP_VAL * TEMP_DELTA)
    Next j
    
    ALPHA_VECTOR(l) = (Log(TEMP_SUM) - Log(HW_PRICE_FUNC(TENOR_RNG, _
                ZERO_RNG, (l + 1) * TEMP_DELTA))) / TEMP_DELTA
    For j = JMINUS_VAL To JPLUS_VAL
        R_ARR(l, j) = ALPHA_VECTOR(l) + j * BTEMP_VAL
        S_ARR(l, j) = HW_CONTINUOUS_RATE_FUNC(R_ARR(l, j), TEMP_DELTA, l * TEMP_DELTA, _
                    SIGMA, KAPPA, TENOR_RNG, ZERO_RNG)
        P_ARR(l, j) = HW_A_FUNC(l * TEMP_DELTA, FIRST_TENOR, SIGMA, KAPPA, _
                    TENOR_RNG, ZERO_RNG) * Exp(-HW_B_FUNC(l * TEMP_DELTA, _
                        FIRST_TENOR, KAPPA) * S_ARR(l, j)) * CASH_FLOW
    Next j
    
    ii = ii - 1
    jj = jj + 1
    
    For j = ii To jj
        If j = jj Then
            Q_ARR(l + 1, j) = Q_ARR(l, j - 1) * HW_PROBABILITIES_TREE_FUNC(1, j - 1, _
                    JMAX_VAL, KAPPA, TEMP_DELTA) * Exp(-R_ARR(l, j - 1) * TEMP_DELTA)
        ElseIf j = jj - 1 Then
            Q_ARR(l + 1, j) = Q_ARR(l, j) * HW_PROBABILITIES_TREE_FUNC(0, j, JMAX_VAL, _
                    KAPPA, TEMP_DELTA) * Exp(-R_ARR(l, j) * TEMP_DELTA) + _
                    Q_ARR(l, j - 1) * HW_PROBABILITIES_TREE_FUNC(1, j - 1, JMAX_VAL, KAPPA, _
                    TEMP_DELTA) * Exp(-R_ARR(l, j - 1) * TEMP_DELTA)
        ElseIf j < jj - 1 And j > ii + 1 Then
            Q_ARR(l + 1, j) = Q_ARR(l, j - 1) * HW_PROBABILITIES_TREE_FUNC(1, j - 1, JMAX_VAL, _
                    KAPPA, TEMP_DELTA) * Exp(-R_ARR(l, j - 1) * TEMP_DELTA) + _
                    Q_ARR(l, j) * HW_PROBABILITIES_TREE_FUNC(0, j, JMAX_VAL, KAPPA, TEMP_DELTA) * _
                    Exp(-R_ARR(l, j) * TEMP_DELTA) + Q_ARR(l, j + 1) * _
                    HW_PROBABILITIES_TREE_FUNC(-1, j + 1, JMAX_VAL, KAPPA, TEMP_DELTA) * _
                    Exp(-R_ARR(l, j + 1) * TEMP_DELTA)
        ElseIf j = ii + 1 Then
            Q_ARR(l + 1, j) = Q_ARR(l, j) * HW_PROBABILITIES_TREE_FUNC(0, j, JMAX_VAL, KAPPA, _
                    TEMP_DELTA) * Exp(-R_ARR(l, j) * TEMP_DELTA) + _
                    Q_ARR(l, j + 1) * HW_PROBABILITIES_TREE_FUNC(-1, j + 1, JMAX_VAL, KAPPA, _
                    TEMP_DELTA) * Exp(-R_ARR(l, j + 1) * TEMP_DELTA)
        ElseIf j = ii Then
            Q_ARR(l + 1, j) = Q_ARR(l, j + 1) * HW_PROBABILITIES_TREE_FUNC(-1, j + 1, _
                    JMAX_VAL, KAPPA, TEMP_DELTA) * Exp(-R_ARR(l, j + 1) * TEMP_DELTA)
        End If
    Next j
Next l
    
For l = NSIZE To nSTEPS
    JPLUS_VAL = NSIZE
    JMINUS_VAL = -NSIZE
    TEMP_SUM = 0
    For j = -NSIZE To NSIZE
        TEMP_SUM = TEMP_SUM + Q_ARR(l, j) * Exp(-j * BTEMP_VAL * TEMP_DELTA)
    Next j
    
    ALPHA_VECTOR(l) = (Log(TEMP_SUM) - Log(HW_PRICE_FUNC(TENOR_RNG, _
            ZERO_RNG, (l + 1) * TEMP_DELTA))) / TEMP_DELTA
    For j = JMINUS_VAL To JPLUS_VAL
        R_ARR(l, j) = ALPHA_VECTOR(l) + j * BTEMP_VAL
        S_ARR(l, j) = HW_CONTINUOUS_RATE_FUNC(R_ARR(l, j), TEMP_DELTA, l * TEMP_DELTA, _
                    SIGMA, KAPPA, TENOR_RNG, ZERO_RNG)
        P_ARR(l, j) = HW_A_FUNC(l * TEMP_DELTA, FIRST_TENOR, SIGMA, KAPPA, _
                    TENOR_RNG, ZERO_RNG) * Exp(-HW_B_FUNC(l * _
                    TEMP_DELTA, FIRST_TENOR, KAPPA) * S_ARR(l, j)) * CASH_FLOW
    Next j
    
    For j = JMINUS_VAL To JPLUS_VAL
    'considering the spezial case when nonstandard branching leads to central
    'nodes with five incoming arrows
        If JMAX_VAL = 2 Then
            Q_ARR(l + 1, 2) = Q_ARR(l, 2) * HW_PROBABILITIES_TREE_FUNC(1, 2, JMAX_VAL, _
                        KAPPA, TEMP_DELTA) * Exp(-R_ARR(l, 2) * TEMP_DELTA) _
                        + Q_ARR(l, 1) * HW_PROBABILITIES_TREE_FUNC(1, 1, JMAX_VAL, KAPPA, _
                        TEMP_DELTA) * Exp(-R_ARR(l, 1) * TEMP_DELTA)
            Q_ARR(l + 1, 1) = Q_ARR(l, 2) * HW_PROBABILITIES_TREE_FUNC(0, 2, JMAX_VAL, KAPPA, _
                        TEMP_DELTA) * Exp(-R_ARR(l, 2) * TEMP_DELTA) + Q_ARR(l, 1) * _
                        HW_PROBABILITIES_TREE_FUNC(0, 1, JMAX_VAL, KAPPA, TEMP_DELTA) * _
                        Exp(-R_ARR(l, 1) * TEMP_DELTA) + Q_ARR(l, 0) * _
                        HW_PROBABILITIES_TREE_FUNC(1, 0, JMAX_VAL, KAPPA, TEMP_DELTA) * _
                        Exp(-R_ARR(l, 0) * TEMP_DELTA)
            Q_ARR(l + 1, 0) = Q_ARR(l, 2) * HW_PROBABILITIES_TREE_FUNC(-1, 2, JMAX_VAL, _
                        KAPPA, TEMP_DELTA) * Exp(-R_ARR(l, 2) * TEMP_DELTA) + _
                        Q_ARR(l, 1) * HW_PROBABILITIES_TREE_FUNC(-1, 1, JMAX_VAL, KAPPA, TEMP_DELTA) * _
                        Exp(-R_ARR(l, 1) * TEMP_DELTA) + Q_ARR(l, 0) * _
                        HW_PROBABILITIES_TREE_FUNC(0, 0, JMAX_VAL, KAPPA, TEMP_DELTA) * _
                        Exp(-R_ARR(l, 0) * TEMP_DELTA) + Q_ARR(l, -1) * _
                        HW_PROBABILITIES_TREE_FUNC(1, -1, JMAX_VAL, KAPPA, TEMP_DELTA) * _
                        Exp(-R_ARR(l, -1) * TEMP_DELTA) + Q_ARR(l, -2) * _
                        HW_PROBABILITIES_TREE_FUNC(1, -2, JMAX_VAL, KAPPA, TEMP_DELTA) * _
                        Exp(-R_ARR(l, -2) * TEMP_DELTA)
            Q_ARR(l + 1, -1) = Q_ARR(l, 0) * HW_PROBABILITIES_TREE_FUNC(-1, 0, JMAX_VAL, KAPPA, _
                        TEMP_DELTA) * Exp(-R_ARR(l, 0) * TEMP_DELTA) + Q_ARR(l, -1) * _
                        HW_PROBABILITIES_TREE_FUNC(0, -1, JMAX_VAL, KAPPA, TEMP_DELTA) * _
                        Exp(-R_ARR(l, -1) * TEMP_DELTA) + Q_ARR(l, -2) * _
                        HW_PROBABILITIES_TREE_FUNC(1, -2, JMAX_VAL, KAPPA, TEMP_DELTA) * _
                        Exp(-R_ARR(l, -2) * TEMP_DELTA)
            Q_ARR(l + 1, -2) = Q_ARR(l, -1) * HW_PROBABILITIES_TREE_FUNC(-1, -1, JMAX_VAL, _
                        KAPPA, TEMP_DELTA) * Exp(-R_ARR(l, -1) * TEMP_DELTA) + _
                        Q_ARR(l, -2) * HW_PROBABILITIES_TREE_FUNC(-1, -2, JMAX_VAL, KAPPA, _
                        TEMP_DELTA) * Exp(-R_ARR(l, -2) * TEMP_DELTA)
        Else
            If j = JPLUS_VAL Then
                Q_ARR(l + 1, j) = Q_ARR(l, j) * HW_PROBABILITIES_TREE_FUNC(1, j, JMAX_VAL, _
                        KAPPA, TEMP_DELTA) * Exp(-R_ARR(l, j) * TEMP_DELTA) + _
                        Q_ARR(l, j - 1) * HW_PROBABILITIES_TREE_FUNC(1, j - 1, JMAX_VAL, KAPPA, _
                        TEMP_DELTA) * Exp(-R_ARR(l, j - 1) * TEMP_DELTA)
            ElseIf j = JPLUS_VAL - 1 Then
                Q_ARR(l + 1, j) = Q_ARR(l, j + 1) * HW_PROBABILITIES_TREE_FUNC(0, j + 1, JMAX_VAL, _
                        KAPPA, TEMP_DELTA) * Exp(-R_ARR(l, j + 1) * TEMP_DELTA) + _
                        Q_ARR(l, j) * HW_PROBABILITIES_TREE_FUNC(0, j, JMAX_VAL, KAPPA, TEMP_DELTA) * _
                        Exp(-R_ARR(l, j) * TEMP_DELTA) + Q_ARR(l, j - 1) * _
                        HW_PROBABILITIES_TREE_FUNC(1, j - 1, JMAX_VAL, KAPPA, TEMP_DELTA) * _
                        Exp(-R_ARR(l, j - 1) * TEMP_DELTA)
            ElseIf j = JPLUS_VAL - 2 Then
                Q_ARR(l + 1, j) = Q_ARR(l, j + 1) * HW_PROBABILITIES_TREE_FUNC(-1, j + 1, _
                        JMAX_VAL, KAPPA, TEMP_DELTA) * Exp(-R_ARR(l, j + 1) * _
                        TEMP_DELTA) + Q_ARR(l, j) * HW_PROBABILITIES_TREE_FUNC(0, j, JMAX_VAL, _
                        KAPPA, TEMP_DELTA) * Exp(-R_ARR(l, j) * TEMP_DELTA) + _
                        Q_ARR(l, j - 1) * HW_PROBABILITIES_TREE_FUNC(1, j - 1, JMAX_VAL, KAPPA, _
                        TEMP_DELTA) * Exp(-R_ARR(l, j - 1) * TEMP_DELTA) + _
                        Q_ARR(l, j + 2) * HW_PROBABILITIES_TREE_FUNC(-1, j + 2, JMAX_VAL, KAPPA, _
                        TEMP_DELTA) * Exp(-R_ARR(l, j + 2) * TEMP_DELTA)
            ElseIf j < JPLUS_VAL - 2 And j > JMINUS_VAL + 2 Then
                Q_ARR(l + 1, j) = Q_ARR(l, j + 1) * HW_PROBABILITIES_TREE_FUNC(-1, j + 1, JMAX_VAL, _
                        KAPPA, TEMP_DELTA) * Exp(-R_ARR(l, j + 1) * TEMP_DELTA) + _
                        Q_ARR(l, j) * HW_PROBABILITIES_TREE_FUNC(0, j, JMAX_VAL, KAPPA, TEMP_DELTA) * _
                        Exp(-R_ARR(l, j) * TEMP_DELTA) + Q_ARR(l, j - 1) * _
                        HW_PROBABILITIES_TREE_FUNC(1, j - 1, JMAX_VAL, KAPPA, TEMP_DELTA) * _
                        Exp(-R_ARR(l, j - 1) * TEMP_DELTA)
            ElseIf j = JMINUS_VAL + 2 Then
                Q_ARR(l + 1, j) = Q_ARR(l, j + 1) * HW_PROBABILITIES_TREE_FUNC(-1, j + 1, JMAX_VAL, _
                        KAPPA, TEMP_DELTA) * Exp(-R_ARR(l, j + 1) * TEMP_DELTA) + _
                        Q_ARR(l, j) * HW_PROBABILITIES_TREE_FUNC(0, j, JMAX_VAL, KAPPA, TEMP_DELTA) * _
                        Exp(-R_ARR(l, j) * TEMP_DELTA) + Q_ARR(l, j - 1) * _
                        HW_PROBABILITIES_TREE_FUNC(1, j - 1, JMAX_VAL, KAPPA, TEMP_DELTA) * _
                        Exp(-R_ARR(l, j - 1) * TEMP_DELTA) + Q_ARR(l, j - 2) * _
                        HW_PROBABILITIES_TREE_FUNC(1, j - 2, JMAX_VAL, KAPPA, TEMP_DELTA) * _
                        Exp(-R_ARR(l, j - 2) * TEMP_DELTA)
            ElseIf j = JMINUS_VAL + 1 Then
                Q_ARR(l + 1, j) = Q_ARR(l, j) * HW_PROBABILITIES_TREE_FUNC(0, j, JMAX_VAL, KAPPA, _
                        TEMP_DELTA) * Exp(-R_ARR(l, j) * TEMP_DELTA) + _
                        Q_ARR(l, j - 1) * HW_PROBABILITIES_TREE_FUNC(0, j - 1, JMAX_VAL, KAPPA, _
                        TEMP_DELTA) * Exp(-R_ARR(l, j - 1) * TEMP_DELTA) + _
                        Q_ARR(l, j + 1) * HW_PROBABILITIES_TREE_FUNC(-1, j + 1, JMAX_VAL, KAPPA, _
                        TEMP_DELTA) * Exp(-R_ARR(l, j + 1) * TEMP_DELTA)
            ElseIf j = JMINUS_VAL Then
                Q_ARR(l + 1, j) = Q_ARR(l, j) * HW_PROBABILITIES_TREE_FUNC(-1, j, JMAX_VAL, KAPPA, _
                        TEMP_DELTA) * Exp(-R_ARR(l, j) * TEMP_DELTA) + Q_ARR(l, j + 1) _
                        * HW_PROBABILITIES_TREE_FUNC(-1, j + 1, JMAX_VAL, KAPPA, TEMP_DELTA) * _
                        Exp(-R_ARR(l, j + 1) * TEMP_DELTA)
            End If
        End If
    Next j
Next l

Select Case OPTION_FLAG
Case 1 ', "c", "call"
    For j = -MINIMUM_FUNC(nSTEPS, JMAX_VAL) To MINIMUM_FUNC(nSTEPS, JMAX_VAL)
        O_ARR(nSTEPS, j) = MINIMUM_FUNC(P_ARR(nSTEPS, j), STRIKE) + O_ARR(nSTEPS, j)
    Next j
    For i = nSTEPS To NSIZE + 1 Step -1 'backward induction
        For j = -NSIZE To NSIZE
            If j = JMAX_VAL Then
                O_ARR(i - 1, j) = MINIMUM_FUNC(STRIKE, (O_ARR(i, j) * _
                        HW_PROBABILITIES_TREE_FUNC(1, j, JMAX_VAL, KAPPA, TEMP_DELTA) + _
                        O_ARR(i, j - 1) * HW_PROBABILITIES_TREE_FUNC(0, j, JMAX_VAL, KAPPA, _
                        TEMP_DELTA) + O_ARR(i, j - 2) * HW_PROBABILITIES_TREE_FUNC(-1, j, _
                        JMAX_VAL, KAPPA, TEMP_DELTA)) * Exp(-R_ARR(i - 1, j) * _
                        TEMP_DELTA)) + O_ARR(i - 1, j)
            ElseIf j < JMAX_VAL And j > JMIN_VAL Then
                O_ARR(i - 1, j) = MINIMUM_FUNC(STRIKE, (O_ARR(i, j + 1) * _
                        HW_PROBABILITIES_TREE_FUNC(1, j, JMAX_VAL, KAPPA, TEMP_DELTA) + _
                        O_ARR(i, j) * HW_PROBABILITIES_TREE_FUNC(0, j, JMAX_VAL, KAPPA, _
                        TEMP_DELTA) + O_ARR(i, j - 1) * HW_PROBABILITIES_TREE_FUNC(-1, j, _
                        JMAX_VAL, KAPPA, TEMP_DELTA)) * Exp(-R_ARR(i - 1, j) * _
                        TEMP_DELTA)) + O_ARR(i - 1, j)
            ElseIf j = JMIN_VAL Then
                O_ARR(i - 1, j) = MINIMUM_FUNC(STRIKE, (O_ARR(i, j) * _
                        HW_PROBABILITIES_TREE_FUNC(-1, j, JMAX_VAL, KAPPA, TEMP_DELTA) + _
                        O_ARR(i, j + 1) * HW_PROBABILITIES_TREE_FUNC(0, j, JMAX_VAL, KAPPA, _
                        TEMP_DELTA) + O_ARR(i, j + 2) * HW_PROBABILITIES_TREE_FUNC(1, j, _
                        JMAX_VAL, KAPPA, TEMP_DELTA)) * Exp(-R_ARR(i - 1, j) * _
                        TEMP_DELTA)) + O_ARR(i - 1, j)
            End If
        Next j
    Next i
    JMINUS_VAL = -NSIZE
    JPLUS_VAL = NSIZE
    For i = NSIZE To 1 Step -1
        JMINUS_VAL = JMINUS_VAL + 1
        JPLUS_VAL = JPLUS_VAL - 1
         For j = JMINUS_VAL To JPLUS_VAL
            O_ARR(i - 1, j) = MINIMUM_FUNC(STRIKE, (O_ARR(i, j + 1) * _
                    HW_PROBABILITIES_TREE_FUNC(1, j, JMAX_VAL, KAPPA, TEMP_DELTA) + _
                    O_ARR(i, j) * HW_PROBABILITIES_TREE_FUNC(0, j, JMAX_VAL, KAPPA, _
                    TEMP_DELTA) + O_ARR(i, j - 1) * HW_PROBABILITIES_TREE_FUNC(-1, j, _
                    JMAX_VAL, KAPPA, TEMP_DELTA)) * Exp(-R_ARR(i - 1, j) * _
                    TEMP_DELTA)) + O_ARR(i - 1, j)
         Next j
    Next i
Case Else
    For j = -MINIMUM_FUNC(nSTEPS, JMAX_VAL) To MINIMUM_FUNC(nSTEPS, JMAX_VAL)
        O_ARR(nSTEPS, j) = MAXIMUM_FUNC(P_ARR(nSTEPS, j), STRIKE) + O_ARR(nSTEPS, j)
    Next j
    'backward induction
    For i = nSTEPS To NSIZE + 1 Step -1
        For j = -NSIZE To NSIZE
            If j = JMAX_VAL Then
                O_ARR(i - 1, j) = MAXIMUM_FUNC(STRIKE, (O_ARR(i, j) * _
                        HW_PROBABILITIES_TREE_FUNC(1, j, JMAX_VAL, KAPPA, TEMP_DELTA) + O_ARR(i, _
                        j - 1) * HW_PROBABILITIES_TREE_FUNC(0, j, JMAX_VAL, KAPPA, TEMP_DELTA) + _
                        O_ARR(i, j - 2) * HW_PROBABILITIES_TREE_FUNC(-1, j, JMAX_VAL, KAPPA, _
                        TEMP_DELTA)) * Exp(-R_ARR(i - 1, j) * TEMP_DELTA)) + _
                        O_ARR(i - 1, j)
            ElseIf j < JMAX_VAL And j > JMIN_VAL Then
                O_ARR(i - 1, j) = MAXIMUM_FUNC(STRIKE, (O_ARR(i, j + 1) * _
                        HW_PROBABILITIES_TREE_FUNC(1, j, JMAX_VAL, KAPPA, TEMP_DELTA) + O_ARR(i, j) _
                        * HW_PROBABILITIES_TREE_FUNC(0, j, JMAX_VAL, KAPPA, TEMP_DELTA) + O_ARR(i, _
                        j - 1) * HW_PROBABILITIES_TREE_FUNC(-1, j, JMAX_VAL, KAPPA, TEMP_DELTA)) * _
                        Exp(-R_ARR(i - 1, j) * TEMP_DELTA)) + O_ARR(i - 1, j)
            ElseIf j = JMIN_VAL Then
                O_ARR(i - 1, j) = MAXIMUM_FUNC(STRIKE, (O_ARR(i, j) * _
                        HW_PROBABILITIES_TREE_FUNC(-1, j, JMAX_VAL, KAPPA, TEMP_DELTA) + _
                        O_ARR(i, j + 1) * HW_PROBABILITIES_TREE_FUNC(0, j, JMAX_VAL, KAPPA, _
                        TEMP_DELTA) + O_ARR(i, j + 2) * HW_PROBABILITIES_TREE_FUNC(1, j, _
                        JMAX_VAL, KAPPA, TEMP_DELTA)) * Exp(-R_ARR(i - 1, j) * _
                        TEMP_DELTA)) + O_ARR(i - 1, j)
            End If
        Next j
    Next i

    JMINUS_VAL = -NSIZE
    JPLUS_VAL = NSIZE

    For i = NSIZE To 1 Step -1
        JMINUS_VAL = JMINUS_VAL + 1
        JPLUS_VAL = JPLUS_VAL - 1
         For j = JMINUS_VAL To JPLUS_VAL
            O_ARR(i - 1, j) = MAXIMUM_FUNC(STRIKE, (O_ARR(i, j + 1) * HW_PROBABILITIES_TREE_FUNC(1, _
                        j, JMAX_VAL, KAPPA, TEMP_DELTA) + O_ARR(i, j) * HW_PROBABILITIES_TREE_FUNC(0, _
                        j, JMAX_VAL, KAPPA, TEMP_DELTA) + O_ARR(i, j - 1) * _
                        HW_PROBABILITIES_TREE_FUNC(-1, j, JMAX_VAL, KAPPA, TEMP_DELTA)) * _
                        Exp(-R_ARR(i - 1, j) * TEMP_DELTA)) + O_ARR(i - 1, j)
         Next j
    Next i
End Select

HW_CALLABLE_FUNC = O_ARR(0, 0)

Exit Function
ERROR_LABEL:
    HW_CALLABLE_FUNC = Err.number
End Function


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : HW_ZERO_COUPON_FUNC
'DESCRIPTION   : Numerical valuation of zero coupon bond options
'LIBRARY       : FIXED_INCOME
'GROUP         : HW_1F
'ID            : 014
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'REFERENCE: http://www.angelfire.com/ny/financeinfo/HullWhite.pdf
'**********************************************************************************
'**********************************************************************************

Function HW_ZERO_COUPON_FUNC(ByVal OPTION_FLAG As Integer, _
ByVal DELTA As Double, _
ByVal SIGMA As Double, _
ByVal KAPPA As Double, _
ByRef FIRST_TENOR_RNG As Variant, _
ByVal SECOND_TENOR As Double, _
ByVal CASH_FLOW As Double, _
ByVal STRIKE As Double, _
ByRef TENOR_RNG As Variant, _
ByRef ZERO_RNG As Variant)

Dim h As Double
Dim i As Double
Dim j As Double
Dim k As Double

Dim ii As Double
Dim jj As Double

Dim JMAX_VAL As Double
Dim JMIN_VAL As Double

Dim JPLUS_VAL As Double
Dim JMINUS_VAL As Double

Dim NSIZE As Double
Dim NROWS As Double
Dim nSTEPS As Double

Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double
Dim CTEMP_VAL As Double
Dim DTEMP_VAL As Double
Dim ETEMP_VAL As Double
Dim FTEMP_VAL As Double

Dim TEMP_SUM As Double

Dim Q_ARR As Variant
Dim O_ARR As Variant
Dim R_ARR As Variant
Dim S_ARR As Variant

Dim ALPHA_VECTOR As Variant
Dim STEPS_VECTOR As Variant
Dim FIRST_TENOR_VECTOR As Variant

On Error GoTo ERROR_LABEL

If IsArray(FIRST_TENOR_RNG) = True Then
    FIRST_TENOR_VECTOR = FIRST_TENOR_RNG
    If UBound(FIRST_TENOR_VECTOR, 1) = 1 Then: _
        FIRST_TENOR_VECTOR = MATRIX_TRANSPOSE_FUNC(FIRST_TENOR_VECTOR)
Else
    ReDim FIRST_TENOR_VECTOR(1 To 1, 1 To 1)
    FIRST_TENOR_VECTOR(1, 1) = FIRST_TENOR_RNG
End If


Select Case OPTION_FLAG
Case -1 '"put", "p", -1
      h = 1
Case Else
      h = -1
End Select

NROWS = UBound(FIRST_TENOR_VECTOR, 1)

If FIRST_TENOR_VECTOR(NROWS, 1) / DELTA - _
    Int(FIRST_TENOR_VECTOR(NROWS, 1) / DELTA) = 0 Then
    nSTEPS = FIRST_TENOR_VECTOR(NROWS, 1) / DELTA
Else
    nSTEPS = Int(FIRST_TENOR_VECTOR(NROWS, 1) / DELTA) + 1
End If

ATEMP_VAL = SIGMA ^ 2 * (1 - Exp(-2 * KAPPA * DELTA)) / (2 * KAPPA)
BTEMP_VAL = Sqr(3 * ATEMP_VAL)

JMAX_VAL = Int(0.184 / (KAPPA * DELTA)) + 1
JMIN_VAL = -JMAX_VAL
NSIZE = MINIMUM_FUNC(nSTEPS, JMAX_VAL)

'redimensioning variables
ReDim ALPHA_VECTOR(0 To nSTEPS)
ReDim Q_ARR(0 To nSTEPS + 1, -NSIZE To NSIZE)
ReDim O_ARR(0 To nSTEPS, -NSIZE To NSIZE)
ReDim R_ARR(0 To nSTEPS, -NSIZE To NSIZE)
ReDim S_ARR(0 To nSTEPS, -NSIZE To NSIZE)
ReDim STEPS_VECTOR(0 To NROWS)

'calculating start values
ALPHA_VECTOR(0) = -Log(HW_PRICE_FUNC(TENOR_RNG, ZERO_RNG, DELTA)) / DELTA
R_ARR(0, 0) = ALPHA_VECTOR(0)
S_ARR(0, 0) = HW_CONTINUOUS_RATE_FUNC(R_ARR(0, 0), DELTA, 0, SIGMA, KAPPA, TENOR_RNG, ZERO_RNG)
Q_ARR(1, 1) = HW_PROBABILITIES_TREE_FUNC(1, 0, JMAX_VAL, KAPPA, DELTA) * Exp(-R_ARR(0, 0) * DELTA)
Q_ARR(1, 0) = HW_PROBABILITIES_TREE_FUNC(0, 0, JMAX_VAL, KAPPA, DELTA) * Exp(-R_ARR(0, 0) * DELTA)
Q_ARR(1, -1) = HW_PROBABILITIES_TREE_FUNC(-1, 0, JMAX_VAL, KAPPA, DELTA) * Exp(-R_ARR(0, 0) * DELTA)

JMINUS_VAL = 0
JPLUS_VAL = 0

ii = -1
jj = 1

For i = 1 To NSIZE - 1
    JMINUS_VAL = JMINUS_VAL - 1
    JPLUS_VAL = JPLUS_VAL + 1
    TEMP_SUM = 0
    For j = JMINUS_VAL To JPLUS_VAL
        TEMP_SUM = TEMP_SUM + Q_ARR(i, j) * Exp(-j * BTEMP_VAL * DELTA)
    Next j
    ALPHA_VECTOR(i) = (Log(TEMP_SUM) - Log(HW_PRICE_FUNC(TENOR_RNG, _
        ZERO_RNG, (i + 1) * DELTA))) / DELTA
    For j = JMINUS_VAL To JPLUS_VAL
        R_ARR(i, j) = ALPHA_VECTOR(i) + j * BTEMP_VAL
        S_ARR(i, j) = HW_CONTINUOUS_RATE_FUNC(R_ARR(i, j), DELTA, i * DELTA, SIGMA, KAPPA, _
            TENOR_RNG, ZERO_RNG)
    Next j
    
    ii = ii - 1
    jj = jj + 1
    
    For j = ii To jj
    
        If j = jj Then
            Q_ARR(i + 1, j) = Q_ARR(i, j - 1) * HW_PROBABILITIES_TREE_FUNC(1, j - 1, _
                            JMAX_VAL, KAPPA, DELTA) * Exp(-R_ARR(i, j - 1) * DELTA)
        ElseIf j = jj - 1 Then
            Q_ARR(i + 1, j) = Q_ARR(i, j) * HW_PROBABILITIES_TREE_FUNC(0, j, JMAX_VAL, _
                            KAPPA, DELTA) * Exp(-R_ARR(i, j) * DELTA) + _
                            Q_ARR(i, j - 1) * HW_PROBABILITIES_TREE_FUNC(1, j - 1, JMAX_VAL, _
                            KAPPA, DELTA) * Exp(-R_ARR(i, j - 1) * DELTA)
        ElseIf j < jj - 1 And j > ii + 1 Then
            Q_ARR(i + 1, j) = Q_ARR(i, j - 1) * HW_PROBABILITIES_TREE_FUNC(1, j - 1, JMAX_VAL, _
                            KAPPA, DELTA) * Exp(-R_ARR(i, j - 1) * DELTA) + _
                            Q_ARR(i, j) * HW_PROBABILITIES_TREE_FUNC(0, j, JMAX_VAL, KAPPA, DELTA) * _
                            Exp(-R_ARR(i, j) * DELTA) + Q_ARR(i, j + 1) * _
                            HW_PROBABILITIES_TREE_FUNC(-1, j + 1, JMAX_VAL, KAPPA, DELTA) * _
                            Exp(-R_ARR(i, j + 1) * DELTA)
        ElseIf j = ii + 1 Then
            Q_ARR(i + 1, j) = Q_ARR(i, j) * HW_PROBABILITIES_TREE_FUNC(0, j, JMAX_VAL, KAPPA, _
                            DELTA) * Exp(-R_ARR(i, j) * DELTA) + Q_ARR(i, j + 1) * _
                            HW_PROBABILITIES_TREE_FUNC(-1, j + 1, JMAX_VAL, KAPPA, DELTA) * _
                            Exp(-R_ARR(i, j + 1) * DELTA)
        ElseIf j = ii Then
            Q_ARR(i + 1, j) = Q_ARR(i, j + 1) * HW_PROBABILITIES_TREE_FUNC(-1, j + 1, JMAX_VAL, _
                            KAPPA, DELTA) * Exp(-R_ARR(i, j + 1) * DELTA)
        End If
    Next j
Next i
    
    'using the Q_ARR(NSIZE) to calculate ALPHA_VECTOR(NSIZE) and so on up to NSTEPS
For i = NSIZE To nSTEPS
    
    JPLUS_VAL = NSIZE
    JMINUS_VAL = -NSIZE
    
    TEMP_SUM = 0
    For j = -NSIZE To NSIZE
    TEMP_SUM = TEMP_SUM + Q_ARR(i, j) * Exp(-j * BTEMP_VAL * DELTA)
    Next j
    
    ALPHA_VECTOR(i) = (Log(TEMP_SUM) - Log(HW_PRICE_FUNC(TENOR_RNG, _
                    ZERO_RNG, (i + 1) * DELTA))) / DELTA
    For j = JMINUS_VAL To JPLUS_VAL
        R_ARR(i, j) = ALPHA_VECTOR(i) + j * BTEMP_VAL
        S_ARR(i, j) = HW_CONTINUOUS_RATE_FUNC(R_ARR(i, j), DELTA, i * DELTA, _
                      SIGMA, KAPPA, TENOR_RNG, ZERO_RNG)
    Next j
    
    For j = JMINUS_VAL To JPLUS_VAL
    'considering the spezial case when nonstandard branching leads to central
    'nodes with five incoming arrows
        If JMAX_VAL = 2 Then
            Q_ARR(i + 1, 2) = Q_ARR(i, 2) * HW_PROBABILITIES_TREE_FUNC(1, 2, _
                            JMAX_VAL, KAPPA, DELTA) * Exp(-R_ARR(i, 2) * _
                            DELTA) + Q_ARR(i, 1) * HW_PROBABILITIES_TREE_FUNC(1, 1, _
                            JMAX_VAL, KAPPA, DELTA) * Exp(-R_ARR(i, 1) * DELTA)
            Q_ARR(i + 1, 1) = Q_ARR(i, 2) * HW_PROBABILITIES_TREE_FUNC(0, 2, JMAX_VAL, _
                            KAPPA, DELTA) * Exp(-R_ARR(i, 2) * DELTA) + _
                            Q_ARR(i, 1) * HW_PROBABILITIES_TREE_FUNC(0, 1, JMAX_VAL, KAPPA, _
                            DELTA) * Exp(-R_ARR(i, 1) * DELTA) + Q_ARR(i, 0) _
                            * HW_PROBABILITIES_TREE_FUNC(1, 0, JMAX_VAL, KAPPA, DELTA) * _
                            Exp(-R_ARR(i, 0) * DELTA)
            Q_ARR(i + 1, 0) = Q_ARR(i, 2) * HW_PROBABILITIES_TREE_FUNC(-1, 2, JMAX_VAL, _
                            KAPPA, DELTA) * Exp(-R_ARR(i, 2) * DELTA) + Q_ARR(i, 1) * _
                            HW_PROBABILITIES_TREE_FUNC(-1, 1, JMAX_VAL, KAPPA, DELTA) * Exp(-R_ARR(i, 1) * _
                            DELTA) + Q_ARR(i, 0) * HW_PROBABILITIES_TREE_FUNC(0, 0, JMAX_VAL, KAPPA, _
                            DELTA) * Exp(-R_ARR(i, 0) * DELTA) + Q_ARR(i, -1) * _
                            HW_PROBABILITIES_TREE_FUNC(1, -1, JMAX_VAL, KAPPA, DELTA) * Exp(-R_ARR(i, -1) _
                            * DELTA) + Q_ARR(i, -2) * HW_PROBABILITIES_TREE_FUNC(1, -2, JMAX_VAL, KAPPA, _
                            DELTA) * Exp(-R_ARR(i, -2) * DELTA)
            Q_ARR(i + 1, -1) = Q_ARR(i, 0) * HW_PROBABILITIES_TREE_FUNC(-1, 0, JMAX_VAL, KAPPA, DELTA) _
                            * Exp(-R_ARR(i, 0) * DELTA) + Q_ARR(i, -1) * _
                            HW_PROBABILITIES_TREE_FUNC(0, -1, JMAX_VAL, KAPPA, DELTA) * Exp(-R_ARR(i, -1) _
                            * DELTA) + Q_ARR(i, -2) * HW_PROBABILITIES_TREE_FUNC(1, -2, JMAX_VAL, KAPPA, _
                            DELTA) * Exp(-R_ARR(i, -2) * DELTA)
            Q_ARR(i + 1, -2) = Q_ARR(i, -1) * HW_PROBABILITIES_TREE_FUNC(-1, -1, JMAX_VAL, KAPPA, DELTA) _
                            * Exp(-R_ARR(i, -1) * DELTA) + Q_ARR(i, -2) * _
                            HW_PROBABILITIES_TREE_FUNC(-1, -2, JMAX_VAL, KAPPA, DELTA) * Exp(-R_ARR(i, -2) _
                            * DELTA)
        Else
            If j = JPLUS_VAL Then
                Q_ARR(i + 1, j) = Q_ARR(i, j) * HW_PROBABILITIES_TREE_FUNC(1, j, JMAX_VAL, _
                                KAPPA, DELTA) * Exp(-R_ARR(i, j) * DELTA) + _
                                Q_ARR(i, j - 1) * HW_PROBABILITIES_TREE_FUNC(1, j - 1, JMAX_VAL, _
                                KAPPA, DELTA) * Exp(-R_ARR(i, j - 1) * DELTA)
            ElseIf j = JPLUS_VAL - 1 Then
                Q_ARR(i + 1, j) = Q_ARR(i, j + 1) * HW_PROBABILITIES_TREE_FUNC(0, j + 1, _
                                JMAX_VAL, KAPPA, DELTA) * Exp(-R_ARR(i, j + 1) * _
                                DELTA) + Q_ARR(i, j) * HW_PROBABILITIES_TREE_FUNC(0, j, _
                                JMAX_VAL, KAPPA, DELTA) * Exp(-R_ARR(i, j) * _
                                DELTA) + Q_ARR(i, j - 1) * HW_PROBABILITIES_TREE_FUNC(1, j - 1, _
                                JMAX_VAL, KAPPA, DELTA) * Exp(-R_ARR(i, j - 1) * DELTA)
            ElseIf j = JPLUS_VAL - 2 Then
                Q_ARR(i + 1, j) = Q_ARR(i, j + 1) * HW_PROBABILITIES_TREE_FUNC(-1, j + 1, JMAX_VAL, _
                                KAPPA, DELTA) * Exp(-R_ARR(i, j + 1) * DELTA) + _
                                Q_ARR(i, j) * HW_PROBABILITIES_TREE_FUNC(0, j, JMAX_VAL, KAPPA, _
                                DELTA) * Exp(-R_ARR(i, j) * DELTA) + Q_ARR(i, j - 1) _
                                * HW_PROBABILITIES_TREE_FUNC(1, j - 1, JMAX_VAL, KAPPA, DELTA) * _
                                Exp(-R_ARR(i, j - 1) * DELTA) + Q_ARR(i, j + 2) * _
                                HW_PROBABILITIES_TREE_FUNC(-1, j + 2, JMAX_VAL, KAPPA, DELTA) * _
                                Exp(-R_ARR(i, j + 2) * DELTA)
            ElseIf j < JPLUS_VAL - 2 And j > JMINUS_VAL + 2 Then
                Q_ARR(i + 1, j) = Q_ARR(i, j + 1) * HW_PROBABILITIES_TREE_FUNC(-1, j + 1, JMAX_VAL, _
                                KAPPA, DELTA) * Exp(-R_ARR(i, j + 1) * DELTA) + _
                                Q_ARR(i, j) * HW_PROBABILITIES_TREE_FUNC(0, j, JMAX_VAL, KAPPA, _
                                DELTA) * Exp(-R_ARR(i, j) * DELTA) + Q_ARR(i, j - 1) _
                                * HW_PROBABILITIES_TREE_FUNC(1, j - 1, JMAX_VAL, KAPPA, DELTA) * _
                                Exp(-R_ARR(i, j - 1) * DELTA)
            ElseIf j = JMINUS_VAL + 2 Then
                Q_ARR(i + 1, j) = Q_ARR(i, j + 1) * HW_PROBABILITIES_TREE_FUNC(-1, j + 1, JMAX_VAL, _
                                KAPPA, DELTA) * Exp(-R_ARR(i, j + 1) * DELTA) + _
                                Q_ARR(i, j) * HW_PROBABILITIES_TREE_FUNC(0, j, JMAX_VAL, KAPPA, DELTA) * _
                                Exp(-R_ARR(i, j) * DELTA) + Q_ARR(i, j - 1) * _
                                HW_PROBABILITIES_TREE_FUNC(1, j - 1, JMAX_VAL, KAPPA, DELTA) * _
                                Exp(-R_ARR(i, j - 1) * DELTA) + Q_ARR(i, j - 2) * _
                                HW_PROBABILITIES_TREE_FUNC(1, j - 2, JMAX_VAL, KAPPA, DELTA) * _
                                Exp(-R_ARR(i, j - 2) * DELTA)
            ElseIf j = JMINUS_VAL + 1 Then
                Q_ARR(i + 1, j) = Q_ARR(i, j) * HW_PROBABILITIES_TREE_FUNC(0, j, JMAX_VAL, KAPPA, _
                                DELTA) * Exp(-R_ARR(i, j) * DELTA) + Q_ARR(i, j - 1) * _
                                HW_PROBABILITIES_TREE_FUNC(0, j - 1, JMAX_VAL, KAPPA, DELTA) * _
                                Exp(-R_ARR(i, j - 1) * DELTA) + Q_ARR(i, j + 1) * _
                                HW_PROBABILITIES_TREE_FUNC(-1, j + 1, JMAX_VAL, KAPPA, DELTA) * _
                                Exp(-R_ARR(i, j + 1) * DELTA)
            ElseIf j = JMINUS_VAL Then
                Q_ARR(i + 1, j) = Q_ARR(i, j) * HW_PROBABILITIES_TREE_FUNC(-1, j, JMAX_VAL, KAPPA, _
                                DELTA) * Exp(-R_ARR(i, j) * DELTA) + Q_ARR(i, j + 1) _
                                * HW_PROBABILITIES_TREE_FUNC(-1, j + 1, JMAX_VAL, KAPPA, DELTA) * _
                                Exp(-R_ARR(i, j + 1) * DELTA)
            End If
        End If
    Next j
    
Next i

'cash flow calculations
For k = 1 To NROWS
    If k = 1 Then
        CTEMP_VAL = FIRST_TENOR_VECTOR(k, 1)
        Else: CTEMP_VAL = FIRST_TENOR_VECTOR(k, 1) - FIRST_TENOR_VECTOR(k - 1, 1)
    End If

    DTEMP_VAL = FIRST_TENOR_VECTOR(k, 1) / DELTA - Int(FIRST_TENOR_VECTOR(k, 1) / DELTA)
    If DTEMP_VAL = 0 Then
        STEPS_VECTOR(k) = FIRST_TENOR_VECTOR(k, 1) / DELTA
        For j = -MINIMUM_FUNC(STEPS_VECTOR(k), JMAX_VAL) To MINIMUM_FUNC(STEPS_VECTOR(k), JMAX_VAL)
            
            O_ARR(0, 0) = MAXIMUM_FUNC(h * (STRIKE - HW_A_FUNC(FIRST_TENOR_VECTOR(k, 1), _
                            SECOND_TENOR, SIGMA, KAPPA, TENOR_RNG, ZERO_RNG) * _
                                    Exp(-HW_B_FUNC(FIRST_TENOR_VECTOR(k, 1), _
                                            SECOND_TENOR, KAPPA) * _
                                                    HW_CONTINUOUS_RATE_FUNC(R_ARR(STEPS_VECTOR(k), j), _
                                                            DELTA, FIRST_TENOR_VECTOR(k, 1), _
                                                                    SIGMA, KAPPA, TENOR_RNG, _
                                                                            ZERO_RNG)) * CASH_FLOW), 0) * _
                                                                                    Q_ARR(STEPS_VECTOR(k), j) + _
                                                                                            O_ARR(0, 0)
        Next j
    Else
        
        STEPS_VECTOR(k) = Int(FIRST_TENOR_VECTOR(k, 1) / DELTA)
        
        For j = -MINIMUM_FUNC(STEPS_VECTOR(k), JMAX_VAL) To MINIMUM_FUNC(STEPS_VECTOR(k), JMAX_VAL)
             ETEMP_VAL = DTEMP_VAL * MAXIMUM_FUNC(h * (STRIKE - _
                    HW_A_FUNC(STEPS_VECTOR(k) * DELTA, SECOND_TENOR, _
                        SIGMA, KAPPA, TENOR_RNG, ZERO_RNG) * _
                                Exp(-HW_B_FUNC(STEPS_VECTOR(k) * DELTA, SECOND_TENOR, _
                                            KAPPA) * HW_CONTINUOUS_RATE_FUNC(R_ARR(STEPS_VECTOR(k), j), _
                                                        DELTA, STEPS_VECTOR(k) * DELTA, SIGMA, KAPPA, _
                                                                TENOR_RNG, ZERO_RNG)) * CASH_FLOW), 0) * _
                                                                        Q_ARR(STEPS_VECTOR(k), j) + ETEMP_VAL
        Next j
        
        For j = -MINIMUM_FUNC(STEPS_VECTOR(k) + 1, JMAX_VAL) To MINIMUM_FUNC(STEPS_VECTOR(k) + 1, JMAX_VAL)
             FTEMP_VAL = (1 - DTEMP_VAL) * MAXIMUM_FUNC(h * (STRIKE - _
                        HW_A_FUNC((STEPS_VECTOR(k) + 1) * DELTA, SECOND_TENOR, _
                                    SIGMA, KAPPA, TENOR_RNG, ZERO_RNG) * _
                                            Exp(-HW_B_FUNC((STEPS_VECTOR(k) + 1) * _
                                                    DELTA, SECOND_TENOR, KAPPA) * _
                                                            HW_CONTINUOUS_RATE_FUNC(R_ARR(STEPS_VECTOR(k) + 1, j), _
                                                                    DELTA, (STEPS_VECTOR(k) + 1) * DELTA, _
                                                                            SIGMA, KAPPA, TENOR_RNG, ZERO_RNG)) * _
                                                                                    CASH_FLOW), 0) * Q_ARR(STEPS_VECTOR(k) _
                                                                                            + 1, j) + FTEMP_VAL
        Next j
        O_ARR(0, 0) = ETEMP_VAL + FTEMP_VAL
    End If
Next k

HW_ZERO_COUPON_FUNC = O_ARR(0, 0)
   
Exit Function
ERROR_LABEL:
    HW_ZERO_COUPON_FUNC = Err.number
End Function


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : HW_COUPON_BOND_OPTION_FUNC
'DESCRIPTION   : Analytical valuation of a European discount bond option
'Jamshidian demonstrated that a European option on a coupon bond can be
'seen as a portfolio of discount bond options. This means that an option
'on a bond with N-coupon payments after option expiary is decomposed into
'N discount bond options.

'As Jamshidian pointed out this decomposition only works in a one factor
'model like the Hull-White model where all rates are perfectly correlated
'to the short rate.
'LIBRARY       : FIXED_INCOME
'GROUP         : HW_1F
'ID            : 015
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'REFERENCE: http://www.angelfire.com/ny/financeinfo/HullWhite.pdf
'**********************************************************************************
'**********************************************************************************

Function HW_COUPON_BOND_OPTION_FUNC(ByVal OPTION_FLAG As Integer, _
ByVal KAPPA As Double, _
ByVal SIGMA As Double, _
ByVal TENOR_RNG As Variant, _
ByVal ZERO_RNG As Variant, _
ByVal STRIKE As Double, _
ByVal FIRST_TENOR As Double, _
ByVal SECOND_TENOR As Double, _
ByVal nSTEPS As Double, _
ByVal COUPON As Double, _
Optional ByVal OUTPUT As Integer = 0)

'j_max =INT(0.184/(a*(t/steps = dt)))+1

Dim i As Double
Dim k As Double

Dim TEMP_SUM As Double
Dim DELTA_COUPON As Double

Dim BOND_VECTOR As Variant
Dim OPTION_VECTOR As Variant

On Error GoTo ERROR_LABEL

ReDim BOND_VECTOR(1 To nSTEPS, 1 To 1)
ReDim OPTION_VECTOR(1 To nSTEPS, 1 To 1)

DELTA_COUPON = (SECOND_TENOR - FIRST_TENOR) / nSTEPS
TEMP_SUM = 0

For i = 1 To nSTEPS
If i = nSTEPS Then k = 100 Else k = 0
    BOND_VECTOR(i, 1) = (COUPON + k) * HW_A_FUNC(FIRST_TENOR, _
        FIRST_TENOR + i * DELTA_COUPON, SIGMA, KAPPA, TENOR_RNG, ZERO_RNG) * _
            Exp(-HW_B_FUNC(FIRST_TENOR, FIRST_TENOR + i * DELTA_COUPON, KAPPA) * _
                HW_BOND_NEWTON_FUNC(0, KAPPA, SIGMA, TENOR_RNG, ZERO_RNG, STRIKE, _
                    FIRST_TENOR, SECOND_TENOR, nSTEPS, COUPON))
    OPTION_VECTOR(i, 1) = HW_EURO_BOND_VALUATION_FUNC(OPTION_FLAG, FIRST_TENOR + i * DELTA_COUPON, _
        COUPON + k, BOND_VECTOR(i, 1), KAPPA, SIGMA, FIRST_TENOR, _
            TENOR_RNG, ZERO_RNG)
    TEMP_SUM = TEMP_SUM + OPTION_VECTOR(i, 1)
Next i

Select Case OUTPUT
    Case 0
        HW_COUPON_BOND_OPTION_FUNC = TEMP_SUM
    Case 1
        HW_COUPON_BOND_OPTION_FUNC = OPTION_VECTOR
    Case Else
        HW_COUPON_BOND_OPTION_FUNC = BOND_VECTOR 'Strike Values
End Select

Exit Function
ERROR_LABEL:
HW_COUPON_BOND_OPTION_FUNC = Err.number
End Function


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : HW_BOND_OBJ_FUNC
'DESCRIPTION   : HW Bond Formula
'LIBRARY       : FIXED_INCOME
'GROUP         : HW_1F
'ID            : 016
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'REFERENCE: http://www.angelfire.com/ny/financeinfo/HullWhite.pdf
'**********************************************************************************
'**********************************************************************************


Private Function HW_BOND_OBJ_FUNC(ByVal SPOT_RATE As Double, _
ByVal KAPPA As Double, _
ByVal SIGMA As Double, _
ByRef TENOR_RNG As Variant, _
ByRef ZERO_RNG As Variant, _
ByVal STRIKE As Double, _
ByVal FIRST_TENOR As Double, _
ByVal SECOND_TENOR As Double, _
ByVal nSTEPS As Double, _
ByVal COUPON As Double)

Dim i As Double
Dim k As Double

Dim TEMP_SUM As Double
Dim DELTA_COUPON As Double

Dim BOND_VECTOR As Variant

On Error GoTo ERROR_LABEL

ReDim BOND_VECTOR(1 To nSTEPS, 1 To 1)

DELTA_COUPON = (SECOND_TENOR - FIRST_TENOR) / nSTEPS
TEMP_SUM = 0

For i = 1 To nSTEPS
    If i = nSTEPS Then k = 100 Else k = 0
        BOND_VECTOR(i, 1) = (COUPON + k) * HW_A_FUNC(FIRST_TENOR, FIRST_TENOR + _
            i * DELTA_COUPON, SIGMA, KAPPA, TENOR_RNG, ZERO_RNG) * _
            Exp(-HW_B_FUNC(FIRST_TENOR, FIRST_TENOR + i * _
            DELTA_COUPON, KAPPA) * SPOT_RATE)

    TEMP_SUM = TEMP_SUM + BOND_VECTOR(i, 1)
Next i
HW_BOND_OBJ_FUNC = TEMP_SUM - STRIKE

Exit Function
ERROR_LABEL:
HW_BOND_OBJ_FUNC = Err.number
End Function

'**********************************************************************************
'**********************************************************************************
'FUNCTION      : HW_BOND_NEWTON_FUNC
'DESCRIPTION   : HW Bond Pricing Newton-Raphson algorithm
'LIBRARY       : FIXED_INCOME
'GROUP         : HW_1F
'ID            : 017
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'REFERENCE: http://www.angelfire.com/ny/financeinfo/HullWhite.pdf
'**********************************************************************************
'**********************************************************************************

Private Function HW_BOND_NEWTON_FUNC(ByVal SPOT_RATE As Double, _
ByVal KAPPA As Double, _
ByVal SIGMA As Double, _
ByRef TENOR_RNG As Variant, _
ByRef ZERO_RNG As Variant, _
ByVal STRIKE As Double, _
ByVal FIRST_TENOR As Double, _
ByVal SECOND_TENOR As Double, _
ByVal nSTEPS As Double, _
ByVal COUPON As Double)

Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double
Dim CTEMP_VAL As Double

Dim epsilon As Double
Dim tolerance As Double

On Error GoTo ERROR_LABEL

epsilon = 0.001
tolerance = 0.0000001

BTEMP_VAL = SPOT_RATE

Do
    ATEMP_VAL = (HW_BOND_OBJ_FUNC(BTEMP_VAL + epsilon, KAPPA, SIGMA, _
           TENOR_RNG, ZERO_RNG, STRIKE, FIRST_TENOR, SECOND_TENOR, nSTEPS, COUPON) - _
           HW_BOND_OBJ_FUNC(BTEMP_VAL - epsilon, KAPPA, SIGMA, TENOR_RNG, _
           ZERO_RNG, STRIKE, FIRST_TENOR, SECOND_TENOR, nSTEPS, COUPON)) / (2 * epsilon)

    BTEMP_VAL = BTEMP_VAL - HW_BOND_OBJ_FUNC(BTEMP_VAL, KAPPA, SIGMA, _
                TENOR_RNG, ZERO_RNG, STRIKE, FIRST_TENOR, SECOND_TENOR, _
                nSTEPS, COUPON) / ATEMP_VAL
    
    CTEMP_VAL = Abs(HW_BOND_OBJ_FUNC(BTEMP_VAL, KAPPA, SIGMA, TENOR_RNG, ZERO_RNG, _
            STRIKE, FIRST_TENOR, SECOND_TENOR, nSTEPS, COUPON) - 0)

Loop Until CTEMP_VAL <= tolerance

HW_BOND_NEWTON_FUNC = BTEMP_VAL

Exit Function
ERROR_LABEL:
HW_BOND_NEWTON_FUNC = Err.number
End Function


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : HW_EURO_BOND_VALUATION_FUNC
'DESCRIPTION   : HW Euro Bond Valuation Function
'LIBRARY       : FIXED_INCOME
'GROUP         : HW_1F
'ID            : 018
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'REFERENCE: http://www.angelfire.com/ny/financeinfo/HullWhite.pdf
'**********************************************************************************
'**********************************************************************************

Function HW_EURO_BOND_VALUATION_FUNC(ByVal OPTION_FLAG As Integer, _
ByVal SECOND_TENOR As Double, _
ByVal CASH_FLOW As Double, _
ByVal STRIKE As Double, _
ByVal KAPPA As Double, _
ByVal SIGMA As Double, _
ByVal FIRST_TENOR As Double, _
ByRef TENOR_RNG As Variant, _
ByRef ZERO_RNG As Variant, _
Optional ByVal CND_TYPE As Integer = 0)

Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double

On Error GoTo ERROR_LABEL

ATEMP_VAL = (SIGMA / KAPPA) * (1 - Exp(-KAPPA * (SECOND_TENOR - FIRST_TENOR))) * _
        Sqr((1 - Exp(-2 * KAPPA * FIRST_TENOR)) / (2 * KAPPA))

BTEMP_VAL = (1 / ATEMP_VAL) * Log((CASH_FLOW * HW_PRICE_FUNC(TENOR_RNG, ZERO_RNG, SECOND_TENOR)) / _
    (HW_PRICE_FUNC(TENOR_RNG, ZERO_RNG, FIRST_TENOR) * STRIKE)) + ATEMP_VAL / 2

Select Case OPTION_FLAG
Case 1 ', "c", "call"
    HW_EURO_BOND_VALUATION_FUNC = CASH_FLOW * HW_PRICE_FUNC(TENOR_RNG, ZERO_RNG, SECOND_TENOR) * _
                    CND_FUNC(BTEMP_VAL, CND_TYPE) - STRIKE * HW_PRICE_FUNC(TENOR_RNG, _
                            ZERO_RNG, FIRST_TENOR) * CND_FUNC(BTEMP_VAL - ATEMP_VAL, CND_TYPE)
Case Else
    HW_EURO_BOND_VALUATION_FUNC = -CASH_FLOW * HW_PRICE_FUNC(TENOR_RNG, ZERO_RNG, SECOND_TENOR) * _
                CND_FUNC(-BTEMP_VAL, CND_TYPE) + STRIKE * HW_PRICE_FUNC(TENOR_RNG, _
                        ZERO_RNG, FIRST_TENOR) * CND_FUNC(-BTEMP_VAL + ATEMP_VAL, CND_TYPE)
End Select

Exit Function
ERROR_LABEL:
HW_EURO_BOND_VALUATION_FUNC = Err.number
End Function


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : HW_CONTINUOUS_RATE_FUNC
'DESCRIPTION   : Transformation of the interest rate in its continuous counterpart
'LIBRARY       : FIXED_INCOME
'GROUP         : HW_1F
'ID            : 019
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'REFERENCE: http://www.angelfire.com/ny/financeinfo/HullWhite.pdf
'**********************************************************************************
'**********************************************************************************

Private Function HW_CONTINUOUS_RATE_FUNC(ByVal ZERO_PRICE As Double, _
ByVal SHORT_SIGMA As Double, _
ByVal REF_TENOR As Double, _
ByVal SIGMA As Double, _
ByVal KAPPA As Double, _
ByRef TENOR_RNG As Variant, _
ByRef ZERO_RNG As Variant)

On Error GoTo ERROR_LABEL

HW_CONTINUOUS_RATE_FUNC = (ZERO_PRICE * SHORT_SIGMA + Log(HW_A_FUNC(REF_TENOR, REF_TENOR + _
                SHORT_SIGMA, SIGMA, KAPPA, TENOR_RNG, ZERO_RNG))) / _
                HW_B_FUNC(REF_TENOR, REF_TENOR + SHORT_SIGMA, KAPPA)
Exit Function
ERROR_LABEL:
HW_CONTINUOUS_RATE_FUNC = Err.number
End Function

'**********************************************************************************
'**********************************************************************************
'FUNCTION      : HW_ZERO_FUNC
'DESCRIPTION   : Calculation of a zero rate given the short rate
'LIBRARY       : FIXED_INCOME
'GROUP         : HW_1F
'ID            : 020
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'REFERENCE: http://www.angelfire.com/ny/financeinfo/HullWhite.pdf
'**********************************************************************************
'**********************************************************************************

Private Function HW_ZERO_FUNC(ByVal SPOT_RATE As Double, _
ByVal START_TENOR As Double, _
ByVal END_TENOR As Double, _
ByVal SIGMA As Double, _
ByVal KAPPA As Double, _
ByRef TENOR_RNG As Variant, _
ByRef ZERO_RNG As Variant)

On Error GoTo ERROR_LABEL

If START_TENOR = END_TENOR Then
    HW_ZERO_FUNC = 0
Else
    HW_ZERO_FUNC = -1 / (END_TENOR - START_TENOR) * Log(HW_A_FUNC(START_TENOR, _
                  END_TENOR, SIGMA, KAPPA, TENOR_RNG, ZERO_RNG)) + 1 / _
                  (END_TENOR - START_TENOR) * HW_B_FUNC(START_TENOR, _
                  END_TENOR, KAPPA) * SPOT_RATE
End If

Exit Function
ERROR_LABEL:
HW_ZERO_FUNC = Err.number
End Function

'**********************************************************************************
'**********************************************************************************
'FUNCTION      : HW_PRICE_FUNC
'DESCRIPTION   : Price of a discount bond at time zero given the initial
'term structure bond without put option
'LIBRARY       : FIXED_INCOME
'GROUP         : HW_1F
'ID            : 021
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'REFERENCE: http://www.angelfire.com/ny/financeinfo/HullWhite.pdf
'**********************************************************************************
'**********************************************************************************

Private Function HW_PRICE_FUNC(ByRef TENOR_RNG As Variant, _
ByRef ZERO_RNG As Variant, _
ByVal REF_TENOR As Double)

On Error GoTo ERROR_LABEL

HW_PRICE_FUNC = Exp(-HW_ZERO_INTERPOLATION_FUNC(TENOR_RNG, ZERO_RNG, REF_TENOR, 0) * REF_TENOR)

Exit Function
ERROR_LABEL:
HW_PRICE_FUNC = Err.number
End Function


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : HW_DISCOUNT_FACTOR_FUNC
'DESCRIPTION   : HW Discount Factor Function
'LIBRARY       : FIXED_INCOME
'GROUP         : HW_1F
'ID            : 022
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'REFERENCE: http://www.angelfire.com/ny/financeinfo/HullWhite.pdf
'**********************************************************************************
'**********************************************************************************

Private Function HW_DISCOUNT_FACTOR_FUNC(ByRef TENORS_RNG As Variant, _
ByRef ZEROS_RNG As Variant, _
ByVal SECOND_TENOR As Double, _
ByVal NROWS As Double)

On Error GoTo ERROR_LABEL
    
HW_DISCOUNT_FACTOR_FUNC = Exp(-1 * _
HW_ZERO_INTERPOLATION_FUNC(TENORS_RNG, ZEROS_RNG, SECOND_TENOR, NROWS) * SECOND_TENOR)
  
Exit Function
ERROR_LABEL:
HW_DISCOUNT_FACTOR_FUNC = Err.number
End Function


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : HW_FORWARD_RATES_FUNC
'DESCRIPTION   : HW FORWARD RATE FUNCTION
'LIBRARY       : FIXED_INCOME
'GROUP         : HW_1F
'ID            : 023
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'REFERENCE: http://www.angelfire.com/ny/financeinfo/HullWhite.pdf
'**********************************************************************************
'**********************************************************************************

Private Function HW_FORWARD_RATES_FUNC(ByRef TENORS_RNG As Variant, _
ByRef ZEROS_RNG As Variant, _
ByVal SECOND_TENOR As Double, _
ByVal DELTA_TENOR As Double, _
ByVal FREQUENCY As Double, _
ByVal NROWS As Double)
    
On Error GoTo ERROR_LABEL
    
HW_FORWARD_RATES_FUNC = (HW_DISCOUNT_FACTOR_FUNC(TENORS_RNG, ZEROS_RNG, SECOND_TENOR, NROWS) - _
                     HW_DISCOUNT_FACTOR_FUNC(TENORS_RNG, ZEROS_RNG, SECOND_TENOR + _
                     DELTA_TENOR, NROWS)) / HW_ANNUITY_FUNC(TENORS_RNG, ZEROS_RNG, _
                     SECOND_TENOR, DELTA_TENOR, FREQUENCY, NROWS)
Exit Function
ERROR_LABEL:
HW_FORWARD_RATES_FUNC = Err.number
End Function


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : HW_ANNUITY_FUNC
'DESCRIPTION   : HW Annuity Function
'LIBRARY       : FIXED_INCOME
'GROUP         : HW_1F
'ID            : 024
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'REFERENCE: http://www.angelfire.com/ny/financeinfo/HullWhite.pdf
'**********************************************************************************
'**********************************************************************************

Private Function HW_ANNUITY_FUNC(ByRef TENORS_RNG As Variant, _
ByRef ZEROS_RNG As Variant, _
ByVal SECOND_TENOR As Double, _
ByVal DELTA_TENOR As Double, _
ByVal FREQUENCY As Double, _
ByVal NROWS As Double)
    
Dim i As Double
Dim nPeriods As Double
Dim ATEMP_SUM As Double
Dim BTEMP_SUM As Double
Dim SPLIT_FACTOR As Double

On Error GoTo ERROR_LABEL

nPeriods = (DELTA_TENOR * FREQUENCY)
SPLIT_FACTOR = 1 - (DELTA_TENOR * FREQUENCY - (DELTA_TENOR * FREQUENCY))
For i = 1 To nPeriods
    ATEMP_SUM = ATEMP_SUM + HW_DISCOUNT_FACTOR_FUNC(TENORS_RNG, _
                ZEROS_RNG, SECOND_TENOR + i / FREQUENCY, NROWS) / FREQUENCY
Next i

BTEMP_SUM = ATEMP_SUM + HW_DISCOUNT_FACTOR_FUNC(TENORS_RNG, ZEROS_RNG, _
            SECOND_TENOR + i / FREQUENCY, NROWS) / FREQUENCY

HW_ANNUITY_FUNC = ATEMP_SUM * SPLIT_FACTOR + BTEMP_SUM * (1 - SPLIT_FACTOR)
  
Exit Function
ERROR_LABEL:
HW_ANNUITY_FUNC = Err.number
End Function

'**********************************************************************************
'**********************************************************************************
'FUNCTION      : HW_A_FUNC
'DESCRIPTION   : HW First Factor Function
'LIBRARY       : FIXED_INCOME
'GROUP         : HW_1F
'ID            : 025
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'REFERENCE: http://www.angelfire.com/ny/financeinfo/HullWhite.pdf
'**********************************************************************************
'**********************************************************************************

Private Function HW_A_FUNC(ByVal FIRST_TENOR As Double, _
ByVal SECOND_TENOR As Double, _
ByVal SIGMA As Double, _
ByVal KAPPA As Double, _
ByRef TENOR_RNG As Variant, _
ByRef ZERO_RNG As Variant)
Dim epsilon  As Double
On Error GoTo ERROR_LABEL
epsilon = 0.0001
HW_A_FUNC = Exp(Log(HW_PRICE_FUNC(TENOR_RNG, ZERO_RNG, SECOND_TENOR) / _
         HW_PRICE_FUNC(TENOR_RNG, ZERO_RNG, FIRST_TENOR)) - _
         HW_B_FUNC(FIRST_TENOR, SECOND_TENOR, KAPPA) * _
         ((Log(HW_PRICE_FUNC(TENOR_RNG, ZERO_RNG, FIRST_TENOR + epsilon)) _
         - Log(HW_PRICE_FUNC(TENOR_RNG, ZERO_RNG, FIRST_TENOR - epsilon))) _
         / (2 * epsilon)) - 1 / (4 * KAPPA ^ 3) * SIGMA ^ 2 * _
         (Exp(-KAPPA * SECOND_TENOR) - Exp(-KAPPA * FIRST_TENOR)) ^ 2 * _
         (Exp(2 * KAPPA * FIRST_TENOR) - 1))
Exit Function
ERROR_LABEL:
HW_A_FUNC = Err.number
End Function

'**********************************************************************************
'**********************************************************************************
'FUNCTION      : HW_B_FUNC
'DESCRIPTION   : HW Second Factor Function
'LIBRARY       : FIXED_INCOME
'GROUP         : HW_1F
'ID            : 026
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'REFERENCE: http://www.angelfire.com/ny/financeinfo/HullWhite.pdf
'**********************************************************************************
'**********************************************************************************

Private Function HW_B_FUNC(ByVal FIRST_TENOR As Double, _
ByVal SECOND_TENOR As Double, _
ByVal KAPPA As Double)
On Error GoTo ERROR_LABEL
HW_B_FUNC = (1 - Exp(-KAPPA * (SECOND_TENOR - FIRST_TENOR))) / KAPPA
Exit Function
ERROR_LABEL:
HW_B_FUNC = Err.number
End Function


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : HW_PROBABILITIES_TREE_FUNC
'DESCRIPTION   : Probabilities in the Hull-White interest rate tree
'LIBRARY       : FIXED_INCOME
'GROUP         : HW_1F
'ID            : 027
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'REFERENCE: http://www.angelfire.com/ny/financeinfo/HullWhite.pdf
'**********************************************************************************
'**********************************************************************************

'// PERFECT

Private Function HW_PROBABILITIES_TREE_FUNC(ByVal TENOR As Double, _
ByVal step As Double, _
ByVal JMAX_VAL As Double, _
ByVal KAPPA As Double, _
ByVal DELTA As Double)

Dim NSIZE As Double

On Error GoTo ERROR_LABEL

NSIZE = Exp(-KAPPA * DELTA) - 1

If Abs(step) < JMAX_VAL Then
    Select Case TENOR
    Case 1
        HW_PROBABILITIES_TREE_FUNC = 1 / 6 + (step ^ 2 * NSIZE ^ 2 + step * NSIZE) / 2
    Case 0
        HW_PROBABILITIES_TREE_FUNC = 2 / 3 - step ^ 2 * NSIZE ^ 2
    Case -1
        HW_PROBABILITIES_TREE_FUNC = 1 / 6 + (step ^ 2 * NSIZE ^ 2 - step * NSIZE) / 2
    End Select
ElseIf step = -JMAX_VAL Then
    Select Case TENOR
    Case 1
        HW_PROBABILITIES_TREE_FUNC = 1 / 6 + (step ^ 2 * NSIZE ^ 2 - step * NSIZE) / 2
    Case 0
        HW_PROBABILITIES_TREE_FUNC = -1 / 3 - step ^ 2 * NSIZE ^ 2 + 2 * step * NSIZE
    Case -1
        HW_PROBABILITIES_TREE_FUNC = 7 / 6 + (step ^ 2 * NSIZE ^ 2 - 3 * step * NSIZE) / 2
    End Select
ElseIf step = JMAX_VAL Then
    Select Case TENOR
    Case 1
        HW_PROBABILITIES_TREE_FUNC = 7 / 6 + (step ^ 2 * NSIZE ^ 2 + 3 * step * NSIZE) / 2
    Case 0
        HW_PROBABILITIES_TREE_FUNC = -1 / 3 - step ^ 2 * NSIZE ^ 2 - 2 * step * NSIZE
    Case -1
        HW_PROBABILITIES_TREE_FUNC = 1 / 6 + (step ^ 2 * NSIZE ^ 2 - step * NSIZE) / 2
    End Select
End If

Exit Function
ERROR_LABEL:
HW_PROBABILITIES_TREE_FUNC = Err.number
End Function


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : HW_ZERO_INTERPOLATION_FUNC
'DESCRIPTION   : HW YIELD INTERPOLATION FUNCTION
'LIBRARY       : FIXED_INCOME
'GROUP         : HW_1F
'ID            : 028
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'REFERENCE: http://www.angelfire.com/ny/financeinfo/HullWhite.pdf
'**********************************************************************************
'**********************************************************************************

Private Function HW_ZERO_INTERPOLATION_FUNC(ByRef TENOR_RNG As Variant, _
ByRef ZERO_RNG As Variant, _
ByVal REF_TENOR As Double, _
Optional ByRef NSIZE As Double = 0)
    
Dim i As Double

Dim ZERO_VECTOR As Variant
Dim TENOR_VECTOR As Variant

On Error GoTo ERROR_LABEL

TENOR_VECTOR = TENOR_RNG
    If UBound(TENOR_VECTOR, 1) = 1 Then: TENOR_VECTOR = MATRIX_TRANSPOSE_FUNC(TENOR_VECTOR)
ZERO_VECTOR = ZERO_RNG
    If UBound(ZERO_VECTOR, 1) = 1 Then: ZERO_VECTOR = MATRIX_TRANSPOSE_FUNC(ZERO_VECTOR)

If UBound(TENOR_VECTOR, 1) <> UBound(ZERO_VECTOR, 1) Then: GoTo ERROR_LABEL

If NSIZE = 0 Then: NSIZE = UBound(TENOR_VECTOR, 1)

i = 0
Do While REF_TENOR >= TENOR_VECTOR(i + 1, 1)
    i = i + 1
    If (i + 1 > NSIZE) Then
        i = NSIZE - 1
        GoTo 1983
    End If
Loop

1983:
If (REF_TENOR < TENOR_VECTOR(1, 1)) Then
    HW_ZERO_INTERPOLATION_FUNC = ZERO_VECTOR(1, 1) + _
                      (ZERO_VECTOR(2, 1) - ZERO_VECTOR(1, 1)) / _
                     (TENOR_VECTOR(2, 1) - TENOR_VECTOR(1, 1)) * _
                     (REF_TENOR - TENOR_VECTOR(1, 1))
Else
    HW_ZERO_INTERPOLATION_FUNC = ZERO_VECTOR(i, 1) + (ZERO_VECTOR(i + 1, 1) - _
                       ZERO_VECTOR(i, 1)) / (TENOR_VECTOR(i + 1, 1) - _
                       TENOR_VECTOR(i, 1)) * (REF_TENOR - TENOR_VECTOR(i, 1))
End If
    

Exit Function
ERROR_LABEL:
HW_ZERO_INTERPOLATION_FUNC = Err.number
End Function

'-----------------------------------------------------------------------------------
'------------------Implementation tools for Hull-White's No-Arbitrage---------------
'-----------------------Implementation of the 1 Factor models-----------------------
'-----------------------------------------------------------------------------------
'----------------------------------------------------------------------------------
'----------------------------------------------------------------------------------
'HW ASSUMPTIONS:
'1.  The market is frictionless in view of no taxes or transaction costs.
'2.  All securities are perfectly divisible.
'3.  The bond market is complete which means that there exists a discount
'bond for each maturity. This is necessary to assure no-arbitrage and to be
'able to find a unique price for each contingent claim either by replicating
'that claim with a combination of other claims or by risk-neutral pricing.
'----------------------------------------------------------------------------------
'----------------------------------------------------------------------------------

'Unlike equilibrium term structure models (e.g., Vasicek, CIR, etc...) which take
'the different stochastic factors as an input and give the term structure (and
'implicitly bond prices) as an output, the no-arbitrage term structure functions
'in this library take the initial term structure as an input by using time-varying
'parameters. This procedure of adjusting parameters so that the initial term
'structure is exactly matched is generally called calibrating.
                    
'----------------------------------------------------------------------------------
'With the 1F no-arbitrage term structure models it is only
'possible to match the term structure of interest rates but not the
'term structure of interest rate volatilities which often can be
'observed in the market. Hence, the second factor in this function
'which determines the term structure, allows the yield curve not only
'to shift up and down but also twist.
'----------------------------------------------------------------------------------


'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
'As it might become obvious now the Hull-White model is a very flexible model
'allowing the user to either use an analytical or numerical solution depending
'on the instrument he wants to price or hedge. This facilitates making consistent
'assumptions about how interest rates evolve. Additionally the mean reverting
'feature of the Hull-White model and its no-arbitrage property are often
'prerequisites of academics and finance professionals.

'However at the time of writing no term structure model has a dominant position
'except the Black-76 model. Therefore most risk management and pricing tools
'give the user a whole set of models to choose from.
'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------

'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
'REFERENCES:
'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------

'Balduzzi, P., S. R. Das, S. Foresi, and R. Sundaram: "A Simple Approach to
'Three-   Factor Affine Term Structure Models", Journal of Fixed Income,
'Dec. 1996,   43-52

'Bjerksund, P., G. Stensland: "Implementation of the Black-Derman-Toy Interest Rate
'Model", Journal of Fixed Income,Sept. 1996, 67-75

'Black, F.: "The Pricing of Commodity Contracts", Journal of Financial Economics,
'March 1976, 167-179

'Black, F., and M. Scholes: "The Pricing of Options and Corporate Liabilities",
'Journal  of Political Economy, 1973, 637-645

'Black, F., E. Derman, and W. Toy: "A One-Factor Model of Interest Rates and Its
'Application to Treasury Bond Options", Financial Analysts Journal, Jan-Feb  1990, 33-39

'Black, F., and P. Karasinski: "Bond and Option Pricing When Short Rates are
'Lognormal", Financial Analyst Journal, July-August, 1991, 52-59

'Brennan, M.J. and E.S. Scwartz: "A Continous Time Approach to Pricing
'Bonds",   Journal of Banking and Finance 3, July 1979, 133-55

'Brennan, M.J. and E.S. Scwartz: "An Equilibrium Model of Bond Pricing and a Test
'of Market Efficiency", Journal of Financial and Quantitative Analysis 17, 3,
'Sep. 1982, 301-29

'Chen, R. R., and l. Scott: "Interest Rate Options in Multifactor Cox-Ingersoll-Ross
'Models of the Term Structure", Journal of Derivatives, Winter 1995, 53-72

'Cox, J.C., J.E. Ingersoll, and S.A. Ross: "An Intertemporal General Equilibrium
'Model of Asset Prices", Econometrica 53, 1985a, 363-384

'Cox, J.C., J.E. Ingersoll, and S.A. Ross: "A Theory of the Term Structure of Interest
'Rates", Econometrica 53, 1985b, 385-407

'Cox, J.C., S.A. Ross and M. Rubinstein: "Option Pricing: A Simplified Approach",
'Journal of Financial Economics 7, Sep. 1979, 229-263

'Duffie, D.: Dynamic Asset Pricing Theory, 2nd edition, Princeton, NJ: Princeton
'University Press, 1996

'Heath, D., R. Jarrow, and A. Morton: "Bond Pricing and the Term Structure of
'Interest Rates: A New Methodology for Contingent Claim Valuation",  Econometrica 60,
'(1992), 77-105

'Ho, T.S.Y. and S.-B. Lee: "Term Structure Movements and Pricing Interest Rate
'Contingent Claims", Journal of Finance 41, (December 1986), 1011-1029

'Ho, T.S.Y.: "Evolution of Interest Rate Models: A Comparison", Journal of Derivatives
'3, 1995, 9-20

'Hull, J. and A. White: "Pricing Interest Rate Derivative Securities, Review of
'Financial Studies 3,4 (1990), 573-592

'Hull, J. and A. White: "One-Factor Interest Rate Models and the Valuation of Interest
'Rate Derivative Securities", Journal of Financial and Quantitative Analysis, 28
'(June 1993a), 235-254

'Hull, J. and A. White: "Bond Option Pricing Based on a Model for the Evolution of
'Bond Prices", Advances in Futures and Options Research, Vol. 6 (1993b) 1-   13

'Hull, J. and A. White: "Branching Out", RISK July 1994a, 34-37

'Hull, J. and A. White: "Numerical Procedures for Implementing Term Structure
'Models I: Single Factor Models", Journal of Derivatives, 2,1 (Fall 1994b), 7-16

'Hull, J. and A. White: "Numerical Procedures for Implementing Term Structure
'Models II: Two-Factor Models", Journal of Derivatives, 2,2 (Winter 1994c), 37-  48

'Hull, J. and A. White: "Using Hull-White Interest Rate Trees", Journal of Derivatives,
'Spring 1996, 26-36

'Hull, J.: Options, Futures, and other Derivatives, 3rd edition,Prentice Hall, 1997

'Jamshidian, F.: "An Exact Bond Option Formula", Journal of Finance, 44, 1989,
'205-  209

'Jamshidian, F.: "Forward Induction and Construction of Yield Curve Diffusion
'Models", Journal of Fixed Income, June 1991, 62-74

'Kraus, A., and M. Smith: "A Simple Multifactor Term Structure Model", Journal of
'Fixed Income, June 1993, 19-23

'Li, A., P. Ritchken, and l. Sankasubramanian: "Lattice Models for Pricing American
'Interest Rate Claims", Journal of Finance, 50, 2 (June 1995), 719-737

'Longstaff, F. A. and E. S. Schwartz: "A Two-Factor Interest Rate Model and
'Contingent Claims Valuation", Journal of Fixed Income, 2, 1992a, 16-23

'Longstaff, F. A. and E. S. Schwartz: "Interest Rate Volatility and the Term Structure:
'A Two-factor General Equilibrium Model", Journal of Finance, 1992b, 1259-   1282

'Musiela, M.,: "General Framework for Pricing Derivative Securities", Stochastic
'Processes and Their Applications, 55, 1995, 227-251

'Rendleman, R. J. jr. and B. J. Bartter: "The Pricing of Options on Debt Securities",
'Journal of Financial and Quantitative Analysis, 15, 1980, 11-24

'Schaefer, S. M. and E. S. Schwartz: "A Two-Factor Model of the Term Structure: An
'Approximation Analytical Solution", Journal of Financial and Quantitative
'Analysis, 19, 1984, 413-448

'Singh, M. K.,: "Estimation of Multifactor Cox, Ingersoll, and Ross Term Structure
'Model: Evidence on Volatility Structure and Parameter Stability", Journal of
'Fixed Income, Sep. 1995, 8-28

'Vasicek, O.: "An Equilibrium Characterisation of the Term Structure", Journal of
'Financial Economics, 5, 1977, 177-188


'--------------------------------------------------------------------------------------
' HW - Key Concepts
'--------------------------------------------------------------------------------------
'1)  Reversion rate of interest rate process
'2)  Value of a money market account
'3)  Price of European call option
'4)  COUPON payment
'5)  Proportional down movement in a binomial model
'6)  Discrete change of the short rate
'7)  Discrete change in time
'8)  Expected value of a variable
'9)  Price of an interest rate derivative
'10) Maximum number of up branches
'11) Tenor of reference rate
'12) Face value of a bond, swap...
'13) Instantaneous drift
'14) Number of time steps
'15) Option value at time step i and state j
'16) Probability of an up movement in a binomial model
'17) Value of a discount bond at time t with maturity   face value 1
'18) State price, Arrow-debreu price in a binomial model
'19) Up-branching probability in Hull-White model
'20) Middle-branching probability in Hull-White model
'21) Down-branching probability in Hull-White model
'22) State price in Hull-White tree
'23) Risk-free rate, short rate
'24) Average short rate
'25) Spot rate at time t with maturity
'26) Strike rate of a derivative
'27) Maturity of a bond, instantaneous standard deviation of the short rate
'28) Maturity of a derivative
'29) Proportional up movement in a binomial model
'30) Strike price of a derivative
'31) Wiener process or Call/Put indicator
'32) Delta of a derivative
'33) Gamma of a derivative
'34) Market price of risk of r
'35) Payout intervall
'--------------------------------------------------------------------------------------
