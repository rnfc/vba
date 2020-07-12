Attribute VB_Name = "FINAN_FI_CDO_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : CDO_SYNTHETIC_PRICE_FUNC
'DESCRIPTION   : Single Tranche Synthetic CDO
'Pricing and Risk Analysis of Correlation Products - Evidence of
'Synthetic CDO Swaps.

'With this function the user just needs to change the Correlation
'and the tranche attachement points to find the corresponding price
'of the tranche.

'LIBRARY       : FIXED_INCOME
'GROUP         : CDO
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'************************************************************************************
'************************************************************************************

'---------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------

'Total Notional:      10,000,000
'Settlement Date:     20/03/2007
'Maturity Date:       20/06/2012

'---------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------
'SYNTHETHIC CDO PRICE EXAMPLE:
'Tranche Name:  equity     Junior1   Mezzanine1  Mezzanine2  Mezzanine3   Senior
'Tranche:        0-3%        3-6%        6-9%       9-12%       12-22%     22-100%
'Market Price:  33.7%      145.2000    74.1000     45.0450     23.8350     13.4150
'Correlation:                           Model quotes
'    14%        47.28%      46.06       44.72       43.44       11.05       0.40
'    20%        50.48%      49.05       49.38       51.48       10.06       0.22
'    25%        54.56%      54.06       55.32       56.23       8.80        0.15
'    30%        59.06%      58.93       59.53       57.97       7.78        0.13
'    40%        65.52%      63.77       60.66       55.93       6.47        0.10
'---------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------

'Reference: Pricing and Risk Analysis of Correlation Products-Evidence of Synthetic CDOs
'http://www.wilmott.com/messageview.cfm?catid=4&threadid=53570


'------------------------------------------------------------------------------------
'Key words: Survival function, joint distribution, loss distribution, Gaussian
'copula, Factor copula, probability bucketing, base correlation, implied correlation.
'------------------------------------------------------------------------------------

Function CDO_SYNTHETIC_PRICE_FUNC(ByVal NOTIONAL As Double, _
ByVal SETTLEMENT As Date, _
ByVal MATURITY As Date, _
ByRef TRANCHES_RNG As Variant, _
ByRef LOWER_K1_RNG As Variant, _
ByRef UPPER_K1_RNG As Variant, _
Optional ByVal FREQUENCY As Integer = 2, _
Optional ByVal FLAT_CORREL As Double = 0.14, _
Optional ByVal FLAT_RECOVER As Double = 0.4, _
Optional ByVal NO_ASSETS As Long = 125, _
Optional ByVal AVG_SPREAD As Double = 54.27, _
Optional ByVal RISK_FREE As Double = 0.05, _
Optional ByVal LOWER_ATTACHMENT_POINTS As Double = 0#, _
Optional ByVal UPPER_ATTACHMENT_POINTS As Double = 0.03, _
Optional ByVal COUNT_BASIS As Integer = 0, _
Optional ByVal OUTPUT As Integer = 0)

'FLAT_CORREL: CORRELATION
'FLAT_RECOVER: FLAT RECOVERY RATE
'FIRST_PAYMENT: NEXT_COUPON_DATE
'FREQUENCY: PAYMENTS_YEAR
'NO_COUPONS: NO_PAYMENTS
'AVG SPREAD --> 5-year Spread

Dim i As Long
Dim j As Long

Dim TEMP1_VAL As Double
Dim TEMP2_VAL As Double
Dim TEMP3_VAL As Double
Dim TEMP4_VAL As Double

Dim TENOR_VECTOR As Variant

Dim TEMP_GROUP As Variant
Dim TEMP_MATRIX As Variant

Dim DISCOUNT_FACTOR As Double
Dim DEFAULT_INTENSITY As Double 'Default Intensity
Dim INTEGRATOR_MATRIX As Variant

Dim NO_COUPONS As Long
Dim FIRST_PAYMENT As Date

Dim NO_TRANCHES As Long
Dim TRANCHES_VECTOR As Variant
Dim LOWER_K1_VECTOR As Variant
Dim UPPER_K1_VECTOR As Variant

ReDim TEMP_GROUP(1 To 8)

On Error GoTo ERROR_LABEL
'---------------------------------------------------------------------------
TRANCHES_VECTOR = TRANCHES_RNG
If UBound(TRANCHES_VECTOR, 1) = 1 Then
    TRANCHES_VECTOR = MATRIX_TRANSPOSE_FUNC(TRANCHES_VECTOR)
End If
NO_TRANCHES = UBound(TRANCHES_VECTOR, 1)
'---------------------------------------------------------------------------
LOWER_K1_VECTOR = LOWER_K1_RNG
If UBound(LOWER_K1_VECTOR, 1) = 1 Then
    LOWER_K1_VECTOR = MATRIX_TRANSPOSE_FUNC(LOWER_K1_VECTOR)
End If
If UBound(LOWER_K1_VECTOR, 1) <> NO_TRANCHES Then: GoTo ERROR_LABEL
'---------------------------------------------------------------------------
UPPER_K1_VECTOR = UPPER_K1_RNG
If UBound(UPPER_K1_VECTOR, 1) = 1 Then
    UPPER_K1_VECTOR = MATRIX_TRANSPOSE_FUNC(UPPER_K1_VECTOR)
End If
If UBound(UPPER_K1_VECTOR, 1) <> NO_TRANCHES Then: GoTo ERROR_LABEL
'-------------------------------------------------------------------------
NO_COUPONS = COUPNUM_FUNC(SETTLEMENT, MATURITY, FREQUENCY)

FIRST_PAYMENT = COUPNCD_FUNC(SETTLEMENT, MATURITY, FREQUENCY)
'-------------------------------------------------------------------------
ReDim TENOR_VECTOR(1 To NO_COUPONS, 1 To 2)
TENOR_VECTOR(1, 1) = FIRST_PAYMENT
TENOR_VECTOR(1, 2) = YEARFRAC_FUNC(SETTLEMENT, TENOR_VECTOR(1, 1), COUNT_BASIS)
For i = 2 To NO_COUPONS
    TENOR_VECTOR(i, 1) = COUPNCD_FUNC(TENOR_VECTOR(i - 1, 1), _
                         MATURITY, FREQUENCY)
    TENOR_VECTOR(i, 2) = YEARFRAC_FUNC(SETTLEMENT, _
                        TENOR_VECTOR(i, 1), COUNT_BASIS)
Next i

'-------------------------------------------------------------------------
'------------------------------FIRST PASS---------------------------------
'-------------------------------------------------------------------------
'Synthetic CDO Structure
'-------------------------------------------------------------------------
'-------------------------------------------------------------------------
ReDim TEMP_MATRIX(0 To NO_TRANCHES + 1, 1 To 7)

TEMP_MATRIX(0, 1) = "TRANCHE_INDEX"
'-------------------------------------------------
'Attachment  Points:
'-------------------------------------------------
TEMP_MATRIX(0, 2) = "K1_LOWER"
TEMP_MATRIX(0, 3) = "K1_UPPER"
'-------------------------------------------------
TEMP_MATRIX(0, 4) = "LOSS_THRESD_LOWER"
TEMP_MATRIX(0, 5) = "LOSS_THRESD_UPPER"
TEMP_MATRIX(0, 6) = "PERCENT"
TEMP_MATRIX(0, 7) = "TRANCHE_NOTIONAL"

TEMP1_VAL = 0
TEMP2_VAL = 0
For i = 1 To NO_TRANCHES
    TEMP_MATRIX(i, 1) = TRANCHES_VECTOR(i, 1)
    TEMP_MATRIX(i, 2) = LOWER_K1_VECTOR(i, 1)
    TEMP_MATRIX(i, 3) = UPPER_K1_VECTOR(i, 1)
    TEMP_MATRIX(i, 4) = NOTIONAL * TEMP_MATRIX(i, 2)
    TEMP_MATRIX(i, 5) = NOTIONAL * TEMP_MATRIX(i, 3)
    TEMP_MATRIX(i, 6) = Abs(TEMP_MATRIX(i, 2) - TEMP_MATRIX(i, 3))
    TEMP1_VAL = TEMP1_VAL + TEMP_MATRIX(i, 6)
    TEMP_MATRIX(i, 7) = TEMP_MATRIX(i, 6) * NOTIONAL
    TEMP2_VAL = TEMP2_VAL + TEMP_MATRIX(i, 7)
Next i

TEMP_MATRIX(NO_TRANCHES + 1, 1) = "TOTALS"
TEMP_MATRIX(NO_TRANCHES + 1, 2) = ""
TEMP_MATRIX(NO_TRANCHES + 1, 3) = ""
TEMP_MATRIX(NO_TRANCHES + 1, 4) = ""
TEMP_MATRIX(NO_TRANCHES + 1, 5) = ""
TEMP_MATRIX(NO_TRANCHES + 1, 6) = TEMP1_VAL
TEMP_MATRIX(NO_TRANCHES + 1, 7) = TEMP2_VAL

'-------------------------------------------------------------------------
TEMP_GROUP(1) = TEMP_MATRIX 'Synthetic CDO Structure
If OUTPUT = 1 Then
    CDO_SYNTHETIC_PRICE_FUNC = TEMP_GROUP(1)
    Exit Function
End If
'-------------------------------------------------------------------------
'-------------------------------------------------------------------------

'-------------------------------------------------------------------------
'------------------------------SECOND PASS---------------------------------
'-------------------------------------------------------------------------
'Portfolio (Loss given Default: 1 - RECOVERY RATE (FLAT_RECOVER))
'-------------------------------------------------------------------------
'-------------------------------------------------------------------------
ReDim TEMP_MATRIX(0 To NO_ASSETS, 1 To 5)

TEMP_MATRIX(0, 1) = ("ASSET (FIRM'S VALUE)")
TEMP_MATRIX(0, 2) = ("NOTIONAL")
TEMP_MATRIX(0, 3) = ("SPREAD")
TEMP_MATRIX(0, 4) = ("RECOVERY RATE")
TEMP_MATRIX(0, 5) = ("LOSS GIVEN DEFAULT")

For i = 1 To NO_ASSETS
    TEMP_MATRIX(i, 1) = i
    TEMP_MATRIX(i, 2) = NOTIONAL / NO_ASSETS
    TEMP_MATRIX(i, 3) = AVG_SPREAD
    TEMP_MATRIX(i, 4) = FLAT_RECOVER
    TEMP_MATRIX(i, 5) = 1 - FLAT_RECOVER
Next i

'-------------------------------------------------------------------------
TEMP_GROUP(2) = TEMP_MATRIX 'Portfolio (Loss given Default)
If OUTPUT = 2 Then 'Loss given Default = 1 - Recovery Rate
    CDO_SYNTHETIC_PRICE_FUNC = TEMP_GROUP(2)
    Exit Function
End If
'-------------------------------------------------------------------------
'-------------------------------------------------------------------------


'-------------------------------------------------------------------------
'------------------------------THIRD PASS---------------------------------
'-------------------------------------------------------------------------
'Cumulative Default Probabilities
'-------------------------------------------------------------------------
'-------------------------------------------------------------------------
ReDim TEMP_MATRIX(0 To NO_ASSETS, 1 To NO_COUPONS)

DEFAULT_INTENSITY = AVG_SPREAD / (1 - FLAT_RECOVER) / 10000
For i = 0 To NO_ASSETS
    If i = 0 Then
        For j = 1 To NO_COUPONS
            TEMP_MATRIX(i, j) = TENOR_VECTOR(j, 1)
        Next j
    Else
        For j = 1 To NO_COUPONS
            TEMP_MATRIX(i, j) = 1 - Exp(-DEFAULT_INTENSITY * _
                                    TENOR_VECTOR(j, 2))
        Next j
    End If
Next i

'-------------------------------------------------------------------------
TEMP_GROUP(3) = TEMP_MATRIX 'Cumulative Default Probabilities
If OUTPUT = 3 Then
    CDO_SYNTHETIC_PRICE_FUNC = TEMP_GROUP(3)
    Exit Function
End If

'-------------------------------------------------------------------------
'-------------------------------------------------------------------------

'-------------------------------------------------------------------------
'------------------------------FORTH PASS---------------------------------
'-------------------------------------------------------------------------
'UnConditional Default Probability
'-------------------------------------------------------------------------
'Unconditional; Portfolio Loss; Conditional; Probability Distribution
'Binomial Distribution
'-------------------------------------------------------------------------
'For the UnConditional Default Probability please check:
'Gauss -Hermite
'Numerical Integration
'http://www.efunda.com/math/num_integration/findgausshermite.cfm

ReDim INTEGRATOR_MATRIX(1 To 4, 1 To 3)

'Common fator
INTEGRATOR_MATRIX(1, 1) = -1.65068012389
INTEGRATOR_MATRIX(2, 1) = -0.524647623275
INTEGRATOR_MATRIX(3, 1) = 0.524647623275
INTEGRATOR_MATRIX(4, 1) = 1.65068012389

'Weights
INTEGRATOR_MATRIX(1, 2) = 1.2402258177
INTEGRATOR_MATRIX(2, 2) = 1.05996448289
INTEGRATOR_MATRIX(3, 2) = 1.05996448289
INTEGRATOR_MATRIX(4, 2) = 1.2402258177

'Integral: SAME AS NORMDIST(FACTOR,0,1,FALSE)*WEIGHT
INTEGRATOR_MATRIX(1, 3) = 0.126689319238926
INTEGRATOR_MATRIX(2, 3) = 0.368494056400546
INTEGRATOR_MATRIX(3, 3) = 0.368494056400546
INTEGRATOR_MATRIX(4, 3) = 0.126689319238926

ReDim TEMP_MATRIX(0 To NO_ASSETS, 1 To NO_COUPONS)

For i = 0 To NO_ASSETS
    If i = 0 Then
        For j = 1 To NO_COUPONS
            TEMP_MATRIX(i, j) = TENOR_VECTOR(j, 1)
        Next j
    Else
'---------------------------------------------------------------------------------
        For j = 1 To NO_COUPONS 'Check This!!!!
'---------------------------------------------------------------------------------
            TEMP1_VAL = (BINOMDIST_FUNC(i, NO_ASSETS, _
                        NORMSDIST_FUNC((NORMSINV_FUNC(TEMP_GROUP(3)(i, j), _
                        0, 1, 1) - Sqr(FLAT_CORREL) * _
                        INTEGRATOR_MATRIX(1, 1)) / Sqr(1 - _
                        FLAT_CORREL), 0, 1, 0), False, _
                        True)) * INTEGRATOR_MATRIX(1, 3)

            TEMP2_VAL = (BINOMDIST_FUNC(i, NO_ASSETS, _
                        NORMSDIST_FUNC((NORMSINV_FUNC(TEMP_GROUP(3)(i, j), _
                        0, 1, 1) - Sqr(FLAT_CORREL) * _
                        INTEGRATOR_MATRIX(2, 1)) / Sqr(1 - _
                        FLAT_CORREL), 0, 1, 0), False, _
                        True)) * INTEGRATOR_MATRIX(2, 3)

            TEMP3_VAL = (BINOMDIST_FUNC(i, NO_ASSETS, _
                        NORMSDIST_FUNC((NORMSINV_FUNC(TEMP_GROUP(3)(i, j), _
                        0, 1, 1) - Sqr(FLAT_CORREL) * _
                        INTEGRATOR_MATRIX(3, 1)) / Sqr(1 - _
                        FLAT_CORREL), 0, 1, 0), False, _
                        True)) * INTEGRATOR_MATRIX(3, 3)

            TEMP4_VAL = (BINOMDIST_FUNC(i, NO_ASSETS, _
                        NORMSDIST_FUNC((NORMSINV_FUNC(TEMP_GROUP(3)(i, j), _
                        0, 1, 1) - Sqr(FLAT_CORREL) * _
                        INTEGRATOR_MATRIX(4, 1)) / Sqr(1 - _
                        FLAT_CORREL), 0, 1, 0), False, _
                        True)) * INTEGRATOR_MATRIX(4, 3)
'-------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------
            TEMP_MATRIX(i, j) = TEMP1_VAL + TEMP2_VAL + TEMP3_VAL + TEMP4_VAL
            'Portfolio Loss Distribution
'-------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------
        Next j
'---------------------------------------------------------------------------------
'---------------------------------------------------------------------------------
    End If
Next i

'-------------------------------------------------------------------------
TEMP_GROUP(4) = TEMP_MATRIX 'UnConditional Default Probability
If OUTPUT = 4 Then
    CDO_SYNTHETIC_PRICE_FUNC = TEMP_GROUP(4)
    Exit Function
End If

'-------------------------------------------------------------------------
'-------------------------------------------------------------------------

'-------------------------------------------------------------------------
'------------------------------FIFTH PASS---------------------------------
'-------------------------------------------------------------------------
'Percentage Loss of the Tranche
'-------------------------------------------------------------------------
'-------------------------------------------------------------------------
ReDim TEMP_MATRIX(0 To NO_ASSETS, 1 To NO_COUPONS)

For i = 0 To NO_ASSETS
    If i = 0 Then
        For j = 1 To NO_COUPONS
            TEMP_MATRIX(i, j) = TENOR_VECTOR(j, 1)
        Next j
    Else
        For j = 1 To NO_COUPONS
            TEMP_MATRIX(i, j) = (MAXIMUM_FUNC(MINIMUM_FUNC(TEMP_GROUP(4)(i, j), _
                                    (UPPER_ATTACHMENT_POINTS)) - _
                                    LOWER_ATTACHMENT_POINTS, 0)) / _
                                    ((UPPER_ATTACHMENT_POINTS) - _
                                    LOWER_ATTACHMENT_POINTS)
        Next j
    End If
Next i

'-------------------------------------------------------------------------
TEMP_GROUP(5) = TEMP_MATRIX 'Cumulative Default Probabilities
If OUTPUT = 5 Then
    CDO_SYNTHETIC_PRICE_FUNC = TEMP_GROUP(5)
    Exit Function
End If

'-------------------------------------------------------------------------
'-------------------------------------------------------------------------

'-------------------------------------------------------------------------
'------------------------------SIXTH PASS---------------------------------
'-------------------------------------------------------------------------
' Expected Tranche Loss
'-------------------------------------------------------------------------
'-------------------------------------------------------------------------
ReDim TEMP_MATRIX(0 To NO_ASSETS + 1, 1 To NO_COUPONS)


For j = 1 To NO_COUPONS
    TEMP_MATRIX(0, j) = TENOR_VECTOR(j, 1)
Next j

For j = 1 To NO_COUPONS
    TEMP1_VAL = 0
    For i = 1 To NO_ASSETS
        TEMP_MATRIX(i, j) = TEMP_GROUP(5)(i, j) * _
                                  TEMP_GROUP(4)(i, j)
        TEMP1_VAL = TEMP1_VAL + TEMP_MATRIX(i, j)
    Next i
    TEMP_MATRIX(NO_ASSETS + 1, j) = TEMP1_VAL
Next j

'-------------------------------------------------------------------------
TEMP_GROUP(6) = TEMP_MATRIX 'Expected Tranche Loss
If OUTPUT = 6 Then
    CDO_SYNTHETIC_PRICE_FUNC = TEMP_GROUP(6)
    Exit Function
End If

'-------------------------------------------------------------------------
'-------------------------------------------------------------------------

'-------------------------------------------------------------------------
'------------------------------SEVENTH PASS-------------------------------
'-------------------------------------------------------------------------
' Default Leg
'-------------------------------------------------------------------------
'-------------------------------------------------------------------------
ReDim TEMP_MATRIX(0 To NO_ASSETS + 1, 1 To NO_COUPONS)

For j = 1 To NO_COUPONS
    TEMP_MATRIX(0, j) = TENOR_VECTOR(j, 1)
Next j

TEMP2_VAL = 0
For j = 1 To NO_COUPONS
    If j = 1 Then
        TEMP1_VAL = 0
        DISCOUNT_FACTOR = (1 / (1 + RISK_FREE) ^ (TENOR_VECTOR(j, 2) / 2))
        For i = 1 To NO_ASSETS
            If i > 2 Then
                TEMP_MATRIX(i, j) = TEMP_MATRIX(i - 2, j) * _
                                    (TEMP_GROUP(6)(i, j) - 0)
            Else
                TEMP_MATRIX(i, j) = DISCOUNT_FACTOR * (TEMP_GROUP(6)(i, j))
            End If
            TEMP1_VAL = TEMP1_VAL + TEMP_MATRIX(i, j)
        Next i
    Else
        TEMP1_VAL = 0
        DISCOUNT_FACTOR = 1 / (1 + RISK_FREE) ^ ((TENOR_VECTOR(j, 2) + _
                        TENOR_VECTOR(j - 1, 2)) / 2)

        For i = 1 To NO_ASSETS
            If TEMP_GROUP(6)(i, j - 1) > TEMP_GROUP(6)(i, j) Then
                TEMP_MATRIX(i, j) = (-1) * DISCOUNT_FACTOR * _
                                    (TEMP_GROUP(6)(i, j) - _
                                     TEMP_GROUP(6)(i, j - 1))
            Else
                TEMP_MATRIX(i, j) = DISCOUNT_FACTOR * (TEMP_GROUP(6)(i, j) - _
                                    TEMP_GROUP(6)(i, j - 1))
            End If

            TEMP1_VAL = TEMP1_VAL + TEMP_MATRIX(i, j)
        Next i
    End If
    TEMP_MATRIX(NO_ASSETS + 1, j) = TEMP1_VAL
    TEMP2_VAL = TEMP2_VAL + TEMP_MATRIX(NO_ASSETS + 1, j)
Next j

'-------------------------------------------------------------------------
TEMP_GROUP(7) = TEMP_MATRIX 'Default Leg
If OUTPUT = 7 Then
    CDO_SYNTHETIC_PRICE_FUNC = TEMP_GROUP(7)
    Exit Function
End If

'-------------------------------------------------------------------------
'-------------------------------------------------------------------------

'-------------------------------------------------------------------------
'------------------------------EIGHT PASS-------------------------------
'-------------------------------------------------------------------------
' Premium Leg
'-------------------------------------------------------------------------
'-------------------------------------------------------------------------
ReDim TEMP_MATRIX(0 To NO_ASSETS + 1, 1 To NO_COUPONS)

For j = 1 To NO_COUPONS
    TEMP_MATRIX(0, j) = TENOR_VECTOR(j, 1)
Next j

TEMP3_VAL = 0
DISCOUNT_FACTOR = 1 / (1 + RISK_FREE) ^ TENOR_VECTOR(1, 2) '-------> CHECK THIS
For j = 1 To NO_COUPONS
    If j = 1 Then
        TEMP1_VAL = 0
        For i = 1 To NO_ASSETS
            'CHECK: DISCOUNT_FACTOR * TENOR_VECTOR(1, 2) '---------> OJO
            TEMP_MATRIX(i, j) = (DISCOUNT_FACTOR * TENOR_VECTOR(1, 2)) _
                            * IIf((UPPER_ATTACHMENT_POINTS - _
                            LOWER_ATTACHMENT_POINTS) < (TEMP_GROUP(6)(i, j) / 2), 0, _
                            (UPPER_ATTACHMENT_POINTS - LOWER_ATTACHMENT_POINTS) _
                            - (TEMP_GROUP(6)(i, j) / 2))
            TEMP1_VAL = TEMP1_VAL + TEMP_MATRIX(i, j)
        Next i
    Else
        TEMP1_VAL = 0
        For i = 1 To NO_ASSETS
            TEMP_MATRIX(i, j) = (DISCOUNT_FACTOR * TENOR_VECTOR(1, 2)) _
                            * IIf((UPPER_ATTACHMENT_POINTS - _
                            LOWER_ATTACHMENT_POINTS) < (TEMP_GROUP(6)(i, j) + _
                            TEMP_GROUP(6)(i, j - 1) / 2), 0, _
                            (UPPER_ATTACHMENT_POINTS - LOWER_ATTACHMENT_POINTS) _
                            - (TEMP_GROUP(6)(i, j) + _
                            TEMP_GROUP(6)(i, j - 1) / 2))
            TEMP1_VAL = TEMP1_VAL + TEMP_MATRIX(i, j)
        Next i
    End If
    TEMP_MATRIX(NO_ASSETS + 1, j) = TEMP1_VAL
    TEMP3_VAL = TEMP3_VAL + TEMP_MATRIX(NO_ASSETS + 1, j)
Next j

'-------------------------------------------------------------------------
TEMP_GROUP(8) = TEMP_MATRIX 'Premium Leg
If OUTPUT = 8 Then
    CDO_SYNTHETIC_PRICE_FUNC = TEMP_GROUP(8)
    Exit Function
End If

'-------------------------------------------------------------------------
'-------------------------------------------------------------------------

CDO_SYNTHETIC_PRICE_FUNC = TEMP2_VAL / TEMP3_VAL 'Market Price

Exit Function
ERROR_LABEL:
CDO_SYNTHETIC_PRICE_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : CDO_GAUSS_SIMULATION_FUNC
'DESCRIPTION   : CDO Pricing with Gaussian Copula
'MC is done without variance reduction
'LIBRARY       : FIXED_INCOME
'GROUP         : CDO
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'************************************************************************************
'************************************************************************************

Function CDO_GAUSS_SIMULATION_FUNC(ByVal NO_ASSETS As Long, _
ByVal NOTION_FIRM As Double, _
ByVal RECOVER_RATE As Double, _
ByVal INTENSITY_CURVE As Double, _
ByVal COPULA_RHO As Double, _
ByVal MATURITY As Double, _
ByVal ATTACH_FIRMS As Double, _
ByVal DETACH_FIRMS As Double, _
ByVal COUPON As Double, _
ByVal FREQUENCY As Double, _
ByVal RATE As Double, _
Optional ByVal ALPHA As Double = 0.05, _
Optional ByVal nLOOPS As Long = 100, _
Optional ByVal PROTEC_BUYER_FLAG As Boolean = True)

'NO_ASSETS: Number of names in the basket
'NOTION_FIRM: Notional per firm (identical for all names)
'RECOVER_RATE: Recovery Rate (identical for all names)
'INTENSITY_CURVE: Forward intensity (identical for all names)
'COPULA_RHO: Base correlation to use in the copula model

'COUPON: Coupon frequency - enter 0.5 for semi annually
'ATTACH_FIRMS: % of Total notional attachement tranche limit
'DETACH_FIRMS: % of Total notional detachement tranche limit
'COUPON: Fixed leg coupon
'RATE: Continuously compound risk-free rate.

Dim i As Long
Dim j As Long
Dim k As Long

Dim TEMP_SUM As Double

Dim LOSS_VAL As Double
Dim FLOAT_VAL As Double
Dim FIXED_VAL As Double
Dim QUANTILE_VAL As Double

Dim MC_CDO_VAL As Double
Dim MC_CBE_VAL As Double
Dim MC_FLOAT_VAL As Double
Dim MC_FIXED_VAL As Double

Dim MC_SQR_CDO_VAL As Double
Dim MC_SQR_CBE_VAL As Double
Dim MC_SQR_FLOAT_VAL As Double
Dim MC_SQR_FIXED_VAL As Double

Dim TEMP_MATRIX As Variant
Dim CORRELATION_MATRIX As Variant
Dim CHOLESKY_MATRIX As Variant
' storing once and for all the cholesky decomposition of the correl matrix
Dim ID_SHOCKS_ARR As Variant
' array to store at each step of the MC the Gaussian IID shocks
Dim CORREL_SHOCKS_ARR As Variant
' array to store at each step of the MC the recorrelated gaussian shocks
Dim DEFAULT_VECTOR As Variant ' default times to store on going
    
On Error GoTo ERROR_LABEL

QUANTILE_VAL = -1 * NORMSINV_FUNC(ALPHA / 2, 0, 1, 0)
ATTACH_FIRMS = ATTACH_FIRMS * NO_ASSETS
DETACH_FIRMS = DETACH_FIRMS * NO_ASSETS

LOSS_VAL = (1 - RECOVER_RATE) 'loss given default per firm

'    setCholeskyMatrix
ReDim CORRELATION_MATRIX(1 To NO_ASSETS, 1 To NO_ASSETS)
For i = 1 To NO_ASSETS
    For j = 1 To NO_ASSETS
        If (i = j) Then
            CORRELATION_MATRIX(i, j) = 1
        Else
            CORRELATION_MATRIX(i, j) = COPULA_RHO
        End If
    Next j
Next i

CHOLESKY_MATRIX = MATRIX_CHOLESKY_FUNC(CORRELATION_MATRIX)


ReDim ID_SHOCKS_ARR(1 To NO_ASSETS)
ReDim CORREL_SHOCKS_ARR(1 To NO_ASSETS)
ReDim DEFAULT_VECTOR(1 To NO_ASSETS, 1 To 1)

MC_FLOAT_VAL = 0
MC_FIXED_VAL = 0
MC_CDO_VAL = 0
MC_CBE_VAL = 0

MC_SQR_FLOAT_VAL = 0
MC_SQR_FIXED_VAL = 0
MC_SQR_CDO_VAL = 0
MC_SQR_CBE_VAL = 0

For k = 1 To nLOOPS
    
    For i = 1 To NO_ASSETS
        ID_SHOCKS_ARR(i) = RANDOM_GAUSS_FUNC(0, 1)
    Next i

    For i = 1 To NO_ASSETS
        TEMP_SUM = 0
        For j = 1 To i 'no need to go further there are zeros in
        '   the upper triandgular
            TEMP_SUM = TEMP_SUM + CHOLESKY_MATRIX(i, j) * ID_SHOCKS_ARR(j)
        Next j
        CORREL_SHOCKS_ARR(i) = TEMP_SUM
    Next i
       
    For i = 1 To NO_ASSETS 'Inverse Dist. Function
        DEFAULT_VECTOR(i, 1) = -Log(1 - (CND_FUNC(CORREL_SHOCKS_ARR(i)))) / INTENSITY_CURVE
    Next i

    DEFAULT_VECTOR = MATRIX_QUICK_SORT_FUNC(DEFAULT_VECTOR, 1, 1) 'sort the
    'default times by ascending order
    FLOAT_VAL = CDO_FLOAT_LEG_FUNC(NO_ASSETS, LOSS_VAL, RATE, FREQUENCY, MATURITY, ATTACH_FIRMS, DETACH_FIRMS, DEFAULT_VECTOR)
    FIXED_VAL = CDO_FIXED_LEG_FUNC(NO_ASSETS, LOSS_VAL, RATE, FREQUENCY, MATURITY, ATTACH_FIRMS, DETACH_FIRMS, DEFAULT_VECTOR)
    MC_FLOAT_VAL = MC_FLOAT_VAL + FLOAT_VAL
    MC_SQR_FLOAT_VAL = MC_SQR_FLOAT_VAL + FLOAT_VAL * FLOAT_VAL
    
    MC_FIXED_VAL = MC_FIXED_VAL + FIXED_VAL
    MC_SQR_FIXED_VAL = MC_SQR_FIXED_VAL + FIXED_VAL * FIXED_VAL
    
    If (PROTEC_BUYER_FLAG) Then
        MC_CDO_VAL = MC_CDO_VAL + (FLOAT_VAL - COUPON * FIXED_VAL)
        MC_SQR_CDO_VAL = MC_SQR_CDO_VAL + (FLOAT_VAL - COUPON * FIXED_VAL) * (FLOAT_VAL - COUPON * FIXED_VAL)
    Else
        MC_CDO_VAL = MC_CDO_VAL + (-FLOAT_VAL + COUPON * FIXED_VAL)
        MC_SQR_CDO_VAL = MC_SQR_CDO_VAL + (-FLOAT_VAL + COUPON * FIXED_VAL) * (-FLOAT_VAL + COUPON * FIXED_VAL)
    End If
        
    MC_CBE_VAL = MC_CBE_VAL + FLOAT_VAL / FIXED_VAL
    MC_SQR_CBE_VAL = MC_SQR_CBE_VAL + (FLOAT_VAL / FIXED_VAL) * (FLOAT_VAL / FIXED_VAL)
Next k

MC_FLOAT_VAL = MC_FLOAT_VAL / nLOOPS
MC_FIXED_VAL = MC_FIXED_VAL / nLOOPS
MC_CDO_VAL = MC_CDO_VAL / nLOOPS
MC_CBE_VAL = MC_CBE_VAL / nLOOPS

MC_SQR_FLOAT_VAL = MC_SQR_FLOAT_VAL / nLOOPS
MC_SQR_FIXED_VAL = MC_SQR_FIXED_VAL / nLOOPS
MC_SQR_CDO_VAL = MC_SQR_CDO_VAL / nLOOPS
MC_SQR_CBE_VAL = MC_SQR_CBE_VAL / nLOOPS

ReDim TEMP_MATRIX(1 To 5, 1 To 5)

TEMP_MATRIX(1, 1) = "SUMMARY"
TEMP_MATRIX(1, 2) = "MC Value"
TEMP_MATRIX(1, 3) = "MC StDev"
TEMP_MATRIX(1, 4) = "LOWER CI"
TEMP_MATRIX(1, 5) = "UPPER CI"

TEMP_MATRIX(2, 1) = "Floating leg value"
TEMP_MATRIX(2, 2) = NOTION_FIRM * MC_FLOAT_VAL
TEMP_MATRIX(2, 3) = NOTION_FIRM * (Sqr(MC_SQR_FLOAT_VAL - MC_FLOAT_VAL * MC_FLOAT_VAL)) / Sqr(nLOOPS)
TEMP_MATRIX(2, 4) = TEMP_MATRIX(2, 2) - QUANTILE_VAL * TEMP_MATRIX(2, 3)
TEMP_MATRIX(2, 5) = TEMP_MATRIX(2, 2) + QUANTILE_VAL * TEMP_MATRIX(2, 3)

TEMP_MATRIX(3, 1) = "Fixed leg value"
TEMP_MATRIX(3, 2) = NOTION_FIRM * MC_FIXED_VAL
TEMP_MATRIX(3, 3) = NOTION_FIRM * (Sqr(MC_SQR_FIXED_VAL - MC_FIXED_VAL * MC_FIXED_VAL)) / Sqr(nLOOPS)
TEMP_MATRIX(3, 4) = TEMP_MATRIX(3, 2) - QUANTILE_VAL * TEMP_MATRIX(3, 3)
TEMP_MATRIX(3, 5) = TEMP_MATRIX(3, 2) + QUANTILE_VAL * TEMP_MATRIX(3, 3)

TEMP_MATRIX(4, 1) = "Total Price with pre defined coupon"
TEMP_MATRIX(4, 2) = NOTION_FIRM * MC_CDO_VAL
TEMP_MATRIX(4, 3) = NOTION_FIRM * (Sqr(MC_SQR_CDO_VAL - MC_CDO_VAL * MC_CDO_VAL)) / Sqr(nLOOPS)
TEMP_MATRIX(4, 4) = TEMP_MATRIX(4, 2) - QUANTILE_VAL * TEMP_MATRIX(4, 3)
TEMP_MATRIX(4, 5) = TEMP_MATRIX(4, 2) + QUANTILE_VAL * TEMP_MATRIX(4, 3)

TEMP_MATRIX(5, 1) = "Break Even coupon (CBE)"
TEMP_MATRIX(5, 2) = MC_CBE_VAL
TEMP_MATRIX(5, 3) = (Sqr(MC_SQR_CBE_VAL - MC_CBE_VAL * MC_CBE_VAL)) / Sqr(nLOOPS)
TEMP_MATRIX(5, 4) = TEMP_MATRIX(5, 2) - QUANTILE_VAL * TEMP_MATRIX(5, 3)
TEMP_MATRIX(5, 5) = TEMP_MATRIX(5, 2) + QUANTILE_VAL * TEMP_MATRIX(5, 3)

CDO_GAUSS_SIMULATION_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
CDO_GAUSS_SIMULATION_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : CDO_FLOAT_LEG_FUNC
'DESCRIPTION   :
'LIBRARY       : FIXED_INCOME
'GROUP         : CDO
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'************************************************************************************
'************************************************************************************

Private Function CDO_FLOAT_LEG_FUNC(ByVal NO_ASSETS As Long, _
ByVal LOSS_VAL As Double, _
ByVal RATE As Double, _
ByVal FREQUENCY As Double, _
ByVal MATURITY As Double, _
ByVal ATTACH_FIRMS As Double, _
ByVal DETACH_FIRMS As Double, _
ByRef DEFAULT_TIME_RNG As Variant)

' each time there is a default, the same constant amount is to be paid,
'ie: loss = NOTION_FIRM * (1 - RECOVER_RATE)
'Hence the only thing we need to do is at each
'DEFAULT_VECTOR <T, pay the discounted lossGivenDefault
    
Dim i As Long

Dim TEMP_SUM As Double
Dim DLOSS_VAL As Double
Dim DTRANCH_VAL As Double 'totalLossesInTheTranch

Dim DEFAULT_VECTOR As Variant

On Error GoTo ERROR_LABEL

DEFAULT_VECTOR = DEFAULT_TIME_RNG
If UBound(DEFAULT_VECTOR, 1) = 1 Then
    DEFAULT_VECTOR = MATRIX_TRANSPOSE_FUNC(DEFAULT_VECTOR)
End If

TEMP_SUM = 0
DTRANCH_VAL = 0
DLOSS_VAL = 0

i = 1
Do While ((i <= NO_ASSETS))
    If ((DEFAULT_VECTOR(i, 1) <= MATURITY) And _
       (DTRANCH_VAL < DETACH_FIRMS - ATTACH_FIRMS)) Then
        DLOSS_VAL = CDO_ACCOUNTS_LOSSES_FUNC(DEFAULT_VECTOR(i, 1), NO_ASSETS, LOSS_VAL, ATTACH_FIRMS, DETACH_FIRMS, DEFAULT_VECTOR) - DTRANCH_VAL
        DTRANCH_VAL = DTRANCH_VAL + DLOSS_VAL
        DLOSS_VAL = DLOSS_VAL - MAXIMUM_FUNC(DTRANCH_VAL - (DETACH_FIRMS - ATTACH_FIRMS), 0)
        TEMP_SUM = TEMP_SUM + Exp(-RATE * DEFAULT_VECTOR(i, 1)) * DLOSS_VAL
        ' assuming that a tranche bears all or nothing of the losses an
        ' extra default this is a good way to see it, as with 100 names,
        ' tranches in percent fit perfectly
    End If
    i = i + 1
Loop

CDO_FLOAT_LEG_FUNC = TEMP_SUM

Exit Function
ERROR_LABEL:
CDO_FLOAT_LEG_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : CDO_FIXED_LEG_FUNC
'DESCRIPTION   :
'LIBRARY       : FIXED_INCOME
'GROUP         : CDO
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'************************************************************************************
'************************************************************************************

Private Function CDO_FIXED_LEG_FUNC(ByVal NO_ASSETS As Long, _
ByVal LOSS_VAL As Double, _
ByVal RATE As Double, _
ByVal FREQUENCY As Double, _
ByVal MATURITY As Double, _
ByVal ATTACH_FIRMS As Double, _
ByVal DETACH_FIRMS As Double, _
ByRef DEFAULT_TIME_RNG As Variant)

' we'll use the course approximation for the FIXED_VAL leg:
' each payment is accounted on
    
Dim i As Long
Dim TEMP_SUM As Double

On Error GoTo ERROR_LABEL

TEMP_SUM = 0
For i = FREQUENCY To MATURITY Step FREQUENCY '0.5 as it is semiannual here
    TEMP_SUM = TEMP_SUM + FREQUENCY * Exp(-RATE * i) * (CDO_OUTSTANDING_NOTIONAL_FUNC(i, NO_ASSETS, LOSS_VAL, ATTACH_FIRMS, DETACH_FIRMS, DEFAULT_TIME_RNG) + CDO_OUTSTANDING_NOTIONAL_FUNC(i - FREQUENCY, NO_ASSETS, LOSS_VAL, ATTACH_FIRMS, DETACH_FIRMS, DEFAULT_TIME_RNG)) / 2
Next i
CDO_FIXED_LEG_FUNC = TEMP_SUM

Exit Function
ERROR_LABEL:
CDO_FIXED_LEG_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : CDO_OUTSTANDING_NOTIONAL_FUNC
'DESCRIPTION   :
'LIBRARY       : FIXED_INCOME
'GROUP         : CDO
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'************************************************************************************
'************************************************************************************

Private Function CDO_OUTSTANDING_NOTIONAL_FUNC(ByVal Y_VAL As Double, _
ByVal NO_ASSETS As Long, _
ByVal LOSS_VAL As Double, _
ByVal ATTACH_FIRMS As Double, _
ByVal DETACH_FIRMS As Double, _
ByRef DEFAULT_TIME_RNG As Variant)
        
On Error GoTo ERROR_LABEL
        
CDO_OUTSTANDING_NOTIONAL_FUNC = ((DETACH_FIRMS - ATTACH_FIRMS) - CDO_ACCOUNTS_LOSSES_FUNC(Y_VAL, NO_ASSETS, LOSS_VAL, ATTACH_FIRMS, DETACH_FIRMS, DEFAULT_TIME_RNG))

Exit Function
ERROR_LABEL:
CDO_OUTSTANDING_NOTIONAL_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CDO_ACCOUNTS_LOSSES_FUNC
'DESCRIPTION   :
'LIBRARY       : FIXED_INCOME
'GROUP         : CDO
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'************************************************************************************
'************************************************************************************

Private Function CDO_ACCOUNTS_LOSSES_FUNC(ByVal Y_VAL As Double, _
ByVal NO_ASSETS As Long, _
ByVal LOSS_VAL As Double, _
ByVal ATTACH_FIRMS As Double, _
ByVal DETACH_FIRMS As Double, _
ByRef DEFAULT_TIME_RNG As Variant)

' as they all have the same notional, the % loss is the # of
' defaults at Y_VAL divided by # firms
    
Dim i As Long
Dim j As Long
Dim DLOSS_VAL As Double
Dim DEFAULT_VECTOR As Variant

On Error GoTo ERROR_LABEL

DEFAULT_VECTOR = DEFAULT_TIME_RNG
If UBound(DEFAULT_VECTOR, 1) = 1 Then
    DEFAULT_VECTOR = MATRIX_TRANSPOSE_FUNC(DEFAULT_VECTOR)
End If

j = 0
i = 1
Do While (i <= NO_ASSETS)
    If ((DEFAULT_VECTOR(i, 1) <= Y_VAL)) Then: j = j + 1
    i = i + 1
Loop
DLOSS_VAL = j * LOSS_VAL
'percent which defaulted times amount to pay = losses
CDO_ACCOUNTS_LOSSES_FUNC = MAXIMUM_FUNC(DLOSS_VAL - ATTACH_FIRMS, 0) - MAXIMUM_FUNC(DLOSS_VAL - DETACH_FIRMS, 0)
Exit Function
ERROR_LABEL:
CDO_ACCOUNTS_LOSSES_FUNC = Err.number
End Function
