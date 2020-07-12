Attribute VB_Name = "FINAN_PORT_FRONTIER_ROY_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'Application of the "Roy Criterion" to the multivariate normal two asset case.

'...we build the efficient frontier and select the frontier portfolio
'which minimizes the probability of not achieving our minimum return...

'In the literature, this procedure is also known as the "Roy Criterion".
'Unfortunately, the simple portfolio construction technique used in this
'illustration is only applicable when returns are (multivariate) normal
'distributed. Contrary to Markowitz portfolio construction, we do not
'make explicit assumptions about the utility function of the investor;
'we assume that his risk preferences are adequatly represented by the
'minimum return and shortfall probability.

'Reference: http://en.wikipedia.org/wiki/Roy's_safety-first_criterion

'Given a portfolio consisting of two assets with the following features...

Function PORT_ROY_CRITERION_ALLOCATION_FUNC( _
ByVal ASSET1_EXPECTED_RETURN_VAL As Double, _
ByVal ASSET1_VOLATILITY_VAL As Double, _
ByVal ASSET2_EXPECTED_RETURN_VAL As Double, _
ByVal ASSET2_VOLATILITY_VAL As Double, _
ByVal ASSET12_RHO_VAL As Double, _
Optional ByRef CASH_RATE_RNG As Variant = 0.02, _
Optional ByVal MIN_WEIGHT_VAL As Double = 0, _
Optional ByVal DELTA_WEIGHT_VAL As Double = 0.05, _
Optional ByVal WNBINS As Double = 21, _
Optional ByVal MIN_VOLATILITY_VAL As Double = 0.255, _
Optional ByVal DELTA_VOLATILITY_VAL As Double = 0.01, _
Optional ByVal VNBINS As Double = 5)

'MIN_RETURN --> certain minimum return that we want to earn...

Dim i As Long
Dim j As Long
Dim k As Long

Dim NROWS As Long

Dim V_VAL As Double
Dim W1_VAL As Double
Dim W2_VAL As Double
Dim EP_VAL As Double
Dim VP_VAL As Double

Dim RATIO1_VAL As Double '(rp-rmin)/vp
Dim RATIO2_VAL As Double '(rp-rmin)/vp
'Dim PROB_VAL As Double 'prob(rp<rmin)

Dim CASH_RATE_VAL As Double

'Dim DATA_MATRIX As Variant
Dim TEMP_MATRIX As Variant
Dim CASH_RATE_VECTOR As Variant

On Error GoTo ERROR_LABEL

'---------------------------------------------------------------------------------------------------------------
If IsArray(CASH_RATE_RNG) = True Then
'---------------------------------------------------------------------------------------------------------------
    CASH_RATE_VECTOR = CASH_RATE_RNG
    If UBound(CASH_RATE_VECTOR, 1) = 1 Then
        CASH_RATE_VECTOR = MATRIX_TRANSPOSE_FUNC(CASH_RATE_VECTOR)
    End If
    NROWS = UBound(CASH_RATE_VECTOR, 1)
    ReDim TEMP_MATRIX(0 To NROWS, 1 To 7 + VNBINS)
    TEMP_MATRIX(0, 1) = "rmin"
    TEMP_MATRIX(0, 2) = "w1"
    TEMP_MATRIX(0, 3) = "w2"
    TEMP_MATRIX(0, 4) = "ep"
    TEMP_MATRIX(0, 5) = "vp"
    TEMP_MATRIX(0, 6) = "max[(rp-rmin)/vp]"
    TEMP_MATRIX(0, 7) = "prob(rp<rmin)"
    
    V_VAL = MIN_VOLATILITY_VAL
    For k = 1 To VNBINS
        TEMP_MATRIX(0, 7 + k) = V_VAL
        V_VAL = V_VAL + DELTA_VOLATILITY_VAL
        'another portfolio with same shortfall prob
    Next k

    For i = 1 To NROWS
        CASH_RATE_VAL = CASH_RATE_VECTOR(i, 1)
        TEMP_MATRIX(i, 1) = CASH_RATE_VAL
        
        RATIO2_VAL = -2 ^ 52
        W1_VAL = MIN_WEIGHT_VAL
        For j = WNBINS To 1 Step -1 'Optimal Portfolios for a given level of cash rate
            W2_VAL = 1 - W1_VAL
            EP_VAL = W1_VAL * ASSET1_EXPECTED_RETURN_VAL + _
                     W2_VAL * ASSET2_EXPECTED_RETURN_VAL
            VP_VAL = Sqr(W1_VAL * W1_VAL * ASSET1_VOLATILITY_VAL * ASSET1_VOLATILITY_VAL + _
                     W2_VAL * W2_VAL * ASSET2_VOLATILITY_VAL * ASSET2_VOLATILITY_VAL + _
                     2 * W1_VAL * W2_VAL * ASSET1_VOLATILITY_VAL * ASSET2_VOLATILITY_VAL _
                     * ASSET12_RHO_VAL)
            If VP_VAL <> 0 Then
                RATIO1_VAL = (EP_VAL - CASH_RATE_VAL) / VP_VAL
                If RATIO1_VAL > RATIO2_VAL Then
                    RATIO2_VAL = RATIO1_VAL
                    TEMP_MATRIX(i, 2) = W1_VAL
                    TEMP_MATRIX(i, 3) = W2_VAL
                    TEMP_MATRIX(i, 4) = EP_VAL
                    TEMP_MATRIX(i, 5) = VP_VAL
                    TEMP_MATRIX(i, 6) = RATIO2_VAL
                    TEMP_MATRIX(i, 7) = NORMSDIST_FUNC((CASH_RATE_VAL - EP_VAL) / VP_VAL, 0, 1, 0)
                    For k = 1 To VNBINS 'another portfolio with same shortfall prob
                        TEMP_MATRIX(i, 7 + k) = TEMP_MATRIX(i, 4) + (TEMP_MATRIX(0, 7 + k) - TEMP_MATRIX(i, 5)) * TEMP_MATRIX(i, 6)
                    Next k
                Else
                    GoTo 1983
                End If
            End If
            W1_VAL = W1_VAL + DELTA_WEIGHT_VAL
        Next j
1983:
    Next i
'---------------------------------------------------------------------------------------------------------------
Else
'---------------------------------------------------------------------------------------------------------------
    ReDim TEMP_MATRIX(0 To WNBINS, 1 To 6)
    TEMP_MATRIX(0, 1) = "w1"
    TEMP_MATRIX(0, 2) = "w2"
    TEMP_MATRIX(0, 3) = "ep"
    TEMP_MATRIX(0, 4) = "vp"
    TEMP_MATRIX(0, 5) = "(rp-rmin)/vp"
    TEMP_MATRIX(0, 6) = "prob(rp<rmin)"
    
    CASH_RATE_VAL = CASH_RATE_RNG
    
    W1_VAL = MIN_WEIGHT_VAL
    For i = WNBINS To 1 Step -1
        TEMP_MATRIX(i, 1) = W1_VAL
        W2_VAL = 1 - W1_VAL
        TEMP_MATRIX(i, 2) = W2_VAL
        TEMP_MATRIX(i, 3) = W1_VAL * ASSET1_EXPECTED_RETURN_VAL + _
                            W2_VAL * ASSET2_EXPECTED_RETURN_VAL
        TEMP_MATRIX(i, 4) = Sqr( _
                W1_VAL * W1_VAL * ASSET1_VOLATILITY_VAL * ASSET1_VOLATILITY_VAL + _
                W2_VAL * W2_VAL * ASSET2_VOLATILITY_VAL * ASSET2_VOLATILITY_VAL + _
                2 * W1_VAL * W2_VAL * ASSET1_VOLATILITY_VAL * ASSET2_VOLATILITY_VAL _
                * ASSET12_RHO_VAL)
        If TEMP_MATRIX(i, 4) <> 0 Then
            TEMP_MATRIX(i, 5) = (TEMP_MATRIX(i, 3) - CASH_RATE_VAL) / TEMP_MATRIX(i, 4)
            TEMP_MATRIX(i, 6) = NORMSDIST_FUNC((CASH_RATE_VAL - TEMP_MATRIX(i, 3)) / TEMP_MATRIX(i, 4), 0, 1, 0)
        End If
        W1_VAL = W1_VAL + DELTA_WEIGHT_VAL
    Next i
'---------------------------------------------------------------------------------------------------------------
End If
'---------------------------------------------------------------------------------------------------------------

PORT_ROY_CRITERION_ALLOCATION_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
PORT_ROY_CRITERION_ALLOCATION_FUNC = Err.number
End Function
