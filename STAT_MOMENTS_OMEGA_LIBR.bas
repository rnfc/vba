Attribute VB_Name = "STAT_MOMENTS_OMEGA_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : VECTOR_OMEGA_FUNC
'DESCRIPTION   :

'We interpret this as follows:
'Suppose that f(x) and F(x) are the probability density and cumulative probability for
'some set of returns. Pick some threshold return r between the Minimum and Maximum returns,
'k and U.

'We consider the numerator in the above expression for Omega:
'1 - F(r) is the probability that a randomly selected return is greater than r.
'(r,U) is the average of returns which are greater than r.
'(r,U) - r is the how much this average exceeds r.
'The numerator is then
'(the probability that a return is greater than r) * (average excess of returns )
'and is a measure of Gains with respect to the selected return r.

'Now, the denominator:
'F(r) is the probability that a randomly selected return is less than r.
'(k,r) is the average of returns which are less than r.
'r - (k,r) is how much r exceeds this average.
'The denominator is then :
'(the probability that a return is less than r) * (average deficit of returns)
'and is a measure of Loss with respect to the selected return r.
'Omega then is a measure of Gains to Losses for the stock in question.
'A stock .. or an entire portfolio of stocks? Either.

'Note that we look at returns above some threshold return r and see if we're in that neighbourhood.
'... as measured by 1 - F(r). Then we look at returns below r and see if we're in that neighbourhood
'... as measured F(r). Since each of these neighbourhood has an associated average return ...

'Reference:
'http://www.gummy-stuff.org/omega2.htm
'http://faculty.fuqua.duke.edu/~charvey/Teaching/BA453_2004/Keating_An_introduction_to.pdf
'http://www.google.ca/search?hl=en&ie=UTF-8&q=%22Omega+ratio%22+%22sharpe+ratio%22&btnG=Search&meta=

'LIBRARY       : STATISTICS
'GROUP         : OMEGA
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 14/02/2008
'************************************************************************************
'************************************************************************************

Function VECTOR_OMEGA_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal TARGET_RATE As Double = 0.03, _
Optional ByVal SLOPE_FACTOR As Double = -0.01, _
Optional ByVal PLOT1_BIN_MIN As Double = -0.15, _
Optional ByVal PLOT1_BIN_WIDTH As Double = 0.003, _
Optional ByVal PLOT2_BIN_MIN As Double = -0.02, _
Optional ByVal PLOT2_BIN_WIDTH As Double = 0.0012, _
Optional ByVal NBINS As Long = 100, _
Optional ByVal DATA_TYPE As Integer = 0, _
Optional ByVal LOG_SCALE As Integer = 0, _
Optional ByVal OUTPUT As Integer = 1)

Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim I1_POINT As Double
Dim I2_POINT As Double
Dim TARGET_POINT As Double
Dim SLOPE_POINT As Double
Dim OMEGA_VAL As Double

Dim DATA_VECTOR As Variant
Dim TEMP_MATRIX As Variant

Dim PLOT1_BIN_MAX As Double
Dim PLOT2_BIN_MAX As Double

Dim PLOT1_ATEMP_SUM As Double
Dim PLOT2_ATEMP_SUM As Double

Dim PLOT1_BTEMP_SUM As Double
Dim PLOT2_BTEMP_SUM As Double

Dim PLOT1_FREQUENCY_VECTOR As Variant
Dim PLOT2_FREQUENCY_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then
    DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
End If

If DATA_TYPE <> 0 Then: DATA_VECTOR = MATRIX_PERCENT_FUNC(DATA_VECTOR, LOG_SCALE)
DATA_VECTOR = MATRIX_QUICK_SORT_FUNC(DATA_VECTOR, 1, 1)
PLOT1_FREQUENCY_VECTOR = HISTOGRAM_FREQUENCY_FUNC(DATA_VECTOR, NBINS, PLOT1_BIN_MIN, PLOT1_BIN_WIDTH, 0)
PLOT1_BIN_MAX = PLOT1_FREQUENCY_VECTOR(NBINS + 1, 1)

PLOT2_FREQUENCY_VECTOR = HISTOGRAM_FREQUENCY_FUNC(DATA_VECTOR, NBINS, PLOT2_BIN_MIN, PLOT2_BIN_WIDTH, 0)
PLOT2_BIN_MAX = PLOT2_FREQUENCY_VECTOR(NBINS + 1, 1)

ReDim TEMP_MATRIX(0 To NBINS + 1, 1 To 16)

TEMP_MATRIX(0, 1) = "PLOT1: BINS"
TEMP_MATRIX(0, 2) = "PLOT1: PROB DENSITY"
TEMP_MATRIX(0, 3) = "PLOT1: CUMUL PROB"
TEMP_MATRIX(0, 4) = "PLOT1: I1"
TEMP_MATRIX(0, 5) = "PLOT1: I2"
TEMP_MATRIX(0, 6) = "PLOT1: I2/I1"

TEMP_MATRIX(0, 7) = "PLOT2: BINS"
TEMP_MATRIX(0, 8) = "PLOT2: PROB DENSITY"
TEMP_MATRIX(0, 9) = "PLOT2: CUMUL PROB"
TEMP_MATRIX(0, 10) = "PLOT2: I1"
TEMP_MATRIX(0, 11) = "PLOT2: I2"
TEMP_MATRIX(0, 12) = "PLOT2: I2/I1"

TEMP_MATRIX(0, 13) = "PLOT1"
TEMP_MATRIX(0, 14) = "PLOT2"

TEMP_MATRIX(0, 15) = "OMEGA SLOPE"
TEMP_MATRIX(0, 16) = "POINT"

'--------------------------------------------------------------------------------------------------
i = 1
TEMP_MATRIX(i, 1) = PLOT1_FREQUENCY_VECTOR(i, 1)
TEMP_MATRIX(i, 2) = PLOT1_FREQUENCY_VECTOR(i, 2)
PLOT1_ATEMP_SUM = TEMP_MATRIX(i, 2)
TEMP_MATRIX(i, 3) = TEMP_MATRIX(i, 2)
TEMP_MATRIX(i, 4) = TEMP_MATRIX(i, 3)

TEMP_MATRIX(i, 7) = PLOT2_FREQUENCY_VECTOR(i, 1)
TEMP_MATRIX(i, 8) = PLOT2_FREQUENCY_VECTOR(i, 2)
PLOT2_ATEMP_SUM = TEMP_MATRIX(i, 2)
TEMP_MATRIX(i, 9) = TEMP_MATRIX(i, 8)
TEMP_MATRIX(i, 10) = TEMP_MATRIX(i, 9)

For i = 2 To NBINS + 1
    TEMP_MATRIX(i, 1) = PLOT1_FREQUENCY_VECTOR(i, 1)
    If (TEMP_MATRIX(i, 1) >= TARGET_RATE) And _
       (TARGET_RATE > TEMP_MATRIX(i - 1, 1)) Then: j = i - 1
    
    TEMP_MATRIX(i, 2) = PLOT1_FREQUENCY_VECTOR(i, 2)
    PLOT1_ATEMP_SUM = PLOT1_ATEMP_SUM + TEMP_MATRIX(i, 2)
    TEMP_MATRIX(i, 3) = TEMP_MATRIX(i - 1, 3) + TEMP_MATRIX(i, 2)
    
    TEMP_MATRIX(i, 7) = PLOT2_FREQUENCY_VECTOR(i, 1)
    If (TEMP_MATRIX(i, 7) >= TARGET_RATE) And _
       (TARGET_RATE > TEMP_MATRIX(i - 1, 7)) Then: k = i - 1

    TEMP_MATRIX(i, 8) = PLOT2_FREQUENCY_VECTOR(i, 2)
    PLOT2_ATEMP_SUM = PLOT2_ATEMP_SUM + TEMP_MATRIX(i, 2)
    TEMP_MATRIX(i, 9) = TEMP_MATRIX(i - 1, 9) + TEMP_MATRIX(i, 8)
Next i

'--------------------------------------------------------------------------------------------------
i = 1
TEMP_MATRIX(i, 3) = TEMP_MATRIX(i, 3) / PLOT1_ATEMP_SUM
PLOT1_BTEMP_SUM = TEMP_MATRIX(i, 3)
TEMP_MATRIX(i, 4) = TEMP_MATRIX(i, 3)

TEMP_MATRIX(i, 9) = TEMP_MATRIX(i, 9) / PLOT2_ATEMP_SUM
PLOT2_BTEMP_SUM = TEMP_MATRIX(i, 9)
TEMP_MATRIX(i, 10) = TEMP_MATRIX(i, 9)

For i = 2 To NBINS + 1
    TEMP_MATRIX(i, 3) = TEMP_MATRIX(i, 3) / PLOT1_ATEMP_SUM
    PLOT1_BTEMP_SUM = PLOT1_BTEMP_SUM + TEMP_MATRIX(i, 3)
    TEMP_MATRIX(i, 4) = TEMP_MATRIX(i - 1, 4) + TEMP_MATRIX(i, 3)
    
    TEMP_MATRIX(i, 9) = TEMP_MATRIX(i, 9) / PLOT2_ATEMP_SUM
    PLOT2_BTEMP_SUM = PLOT2_BTEMP_SUM + TEMP_MATRIX(i, 9)
    TEMP_MATRIX(i, 10) = TEMP_MATRIX(i - 1, 10) + TEMP_MATRIX(i, 9)
Next i

PLOT1_BTEMP_SUM = PLOT1_BTEMP_SUM * PLOT1_BIN_WIDTH
PLOT2_BTEMP_SUM = PLOT2_BTEMP_SUM * PLOT2_BIN_WIDTH

'--------------------------------------------------------------------------------------------------
i = 1
TEMP_MATRIX(i, 4) = TEMP_MATRIX(i, 4) * PLOT1_BIN_WIDTH
TEMP_MATRIX(i, 5) = PLOT1_BIN_MAX - TEMP_MATRIX(i, 1) - (PLOT1_BTEMP_SUM - TEMP_MATRIX(i, 4))
If TEMP_MATRIX(i, 4) <> 0 Then
    TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 5) / TEMP_MATRIX(i, 4)
Else
    TEMP_MATRIX(i, 6) = 0
End If
TEMP_MATRIX(i, 10) = TEMP_MATRIX(i, 10) * PLOT2_BIN_WIDTH
TEMP_MATRIX(i, 11) = PLOT2_BIN_MAX - TEMP_MATRIX(i, 7) - (PLOT2_BTEMP_SUM - TEMP_MATRIX(i, 10))

If TEMP_MATRIX(i, 10) <> 0 Then
    TEMP_MATRIX(i, 12) = TEMP_MATRIX(i, 11) / TEMP_MATRIX(i, 10)
Else
    TEMP_MATRIX(i, 12) = 0
End If
TEMP_MATRIX(i, 13) = IIf(TEMP_MATRIX(i, 1) < TARGET_RATE, TEMP_MATRIX(i, 3), 0)
TEMP_MATRIX(i, 14) = IIf(TEMP_MATRIX(i, 1) < TARGET_RATE, 0, 1 - TEMP_MATRIX(i, 3))
TEMP_MATRIX(i, 15) = 0
TEMP_MATRIX(i, 16) = ""

For i = 2 To NBINS + 1
    TEMP_MATRIX(i, 4) = TEMP_MATRIX(i, 4) * PLOT1_BIN_WIDTH
    TEMP_MATRIX(i, 5) = PLOT1_BIN_MAX - TEMP_MATRIX(i, 1) - (PLOT1_BTEMP_SUM - TEMP_MATRIX(i, 4))
    If TEMP_MATRIX(i, 4) <> 0 Then
        TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 5) / TEMP_MATRIX(i, 4)
    Else
        TEMP_MATRIX(i, 6) = 0
    End If
    TEMP_MATRIX(i, 10) = TEMP_MATRIX(i, 10) * PLOT2_BIN_WIDTH
    TEMP_MATRIX(i, 11) = PLOT2_BIN_MAX - TEMP_MATRIX(i, 7) - (PLOT2_BTEMP_SUM - TEMP_MATRIX(i, 10))
    If TEMP_MATRIX(i, 10) <> 0 Then
        TEMP_MATRIX(i, 12) = TEMP_MATRIX(i, 11) / TEMP_MATRIX(i, 10)
    Else
        TEMP_MATRIX(i, 12) = 0
    End If
    TEMP_MATRIX(i, 13) = IIf(TEMP_MATRIX(i, 1) < TARGET_RATE, TEMP_MATRIX(i, 3), 0)
    TEMP_MATRIX(i, 14) = IIf(TEMP_MATRIX(i, 1) < TARGET_RATE, 0, 1 - TEMP_MATRIX(i, 3))
    TEMP_MATRIX(i, 15) = (TEMP_MATRIX(i, 12) - TEMP_MATRIX(i - 1, 12)) / (TEMP_MATRIX(i, 7) - TEMP_MATRIX(i - 1, 7))
    If TEMP_MATRIX(i - 1, 7) < SLOPE_FACTOR And TEMP_MATRIX(i, 7) > SLOPE_FACTOR Then
        TEMP_MATRIX(i, 16) = TEMP_MATRIX(i - 1, 15) + TEMP_MATRIX(i, 15) * (SLOPE_FACTOR - TEMP_MATRIX(i - 1, 7))
        l = i
    Else
        TEMP_MATRIX(i, 16) = ""
    End If
Next i

Select Case OUTPUT
Case 0
    VECTOR_OMEGA_FUNC = TEMP_MATRIX
Case Else
    
    If j = 0 Then: GoTo 1982
    I1_POINT = TEMP_MATRIX(j, 3)
    I2_POINT = PLOT1_BIN_MAX - TARGET_RATE - TEMP_MATRIX(j, 4)
    If I1_POINT <> 0 Then
        OMEGA_VAL = I2_POINT / I1_POINT
    Else
        OMEGA_VAL = 0
    End If
1982:
    If k = 0 Then: GoTo 1983
    TARGET_POINT = TEMP_MATRIX(k, 12) 'Slope at this point
1983:
    If l = 0 Then: GoTo 1984
    SLOPE_POINT = TEMP_MATRIX(l, 16) 'Slope at this point
1984:
    
    If OUTPUT = 1 Then
        VECTOR_OMEGA_FUNC = Array(OMEGA_VAL, I1_POINT, I2_POINT, TARGET_POINT, SLOPE_POINT)
    Else
        VECTOR_OMEGA_FUNC = Array(OMEGA_VAL, I1_POINT, I2_POINT, TARGET_POINT, SLOPE_POINT, TEMP_MATRIX)
    End If
End Select

Exit Function
ERROR_LABEL:
VECTOR_OMEGA_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : VECTOR_OMEGA_GAMMA_FUNC
'DESCRIPTION   :

'In a recent paper (in PDF format), Keating and Chadwick introduce an
'interesting measure of stock return distributions called Omega.

'It compares the average value of returns above some threshold return r
'(such as a Risk-free return) with the returns less than r.

'While everyone knows that mean and variance cannot capture all of the risk and
'reward features in a financial returns distribution, except in the case where
'returns are normally distributed, performance measurement traditionally relies
'on tools which are based on mean and variance. This has been a matter of
'practicality as econometric attempts to incorporate higher moment effects
'suffer both from complexity of added assumptions and apparently insuperable
'difficulties in their calibration and application due to sparse and noisy data.

'A measure, known as Omega, which employs all the information contained within the
'returns series was introduced in a recent paper. It can be used to rank and evaluate
'portfolios unequivocally. All that is known about the risk and return of a portfolio is
'contained within this measure. With tongue in cheek, it might be considered a
'Sharpe ratio, or the successor to Jensen’s alpha.

'The approach is based upon new insights and developments in mathematical
'techniques, which facilitate the analysis of (returns) distributions.
'In the simplest of terms, it involves partitioning returns into loss and gain
'above and below a return threshold and then considering the probability weighted
'ratio of returns above and below the partitioning.

'For the Sharpe Ratio we look at historical returns ignoring the distribution of
'returns. On the other hand, to evaluate Omega you need to investigate the entire
'distribution of returns.

'REFERENCE:

'http://faculty.fuqua.duke.edu/~charvey/Teaching/BA453_
'2004/Keating_An_introduction_to.pdf

'LIBRARY       : STATISTICS
'GROUP         : OMEGA
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 14/02/2008
'************************************************************************************
'************************************************************************************

Function VECTOR_OMEGA_GAMMA_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal TARGET_RATE As Double = 1, _
Optional ByVal DATA_TYPE As Integer = 0, _
Optional ByVal LOG_SCALE As Integer = 0, _
Optional ByVal OUTPUT As Integer = 0)

Dim i As Long
Dim j As Long

Dim NROWS As Long

Dim BIN_MIN As Double
Dim BIN_WIDTH As Double
Dim NBINS As Double

Dim MIN_VAL As Double
Dim MAX_VAL As Double

Dim TEMP_SUM As Double
Dim TEMP_VAL As Double

Dim TEMP_MATRIX As Variant
Dim DATA_VECTOR As Variant
Dim SUMMARY_VECTOR As Variant
Dim FREQUENCY_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then
    DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
End If
If DATA_TYPE <> 0 Then: DATA_VECTOR = MATRIX_PERCENT_FUNC(DATA_VECTOR, LOG_SCALE)

NROWS = UBound(DATA_VECTOR, 1)

MIN_VAL = 2 ^ 52: MAX_VAL = -2 ^ 52
For i = 1 To NROWS
    If DATA_VECTOR(i, 1) < MIN_VAL Then: MIN_VAL = DATA_VECTOR(i, 1)
    If DATA_VECTOR(i, 1) > MAX_VAL Then: MAX_VAL = DATA_VECTOR(i, 1)
Next i

FREQUENCY_VECTOR = HISTOGRAM_BIN_LIMITS_FUNC(MIN_VAL, MAX_VAL, NROWS, 3)
BIN_WIDTH = FREQUENCY_VECTOR(LBound(FREQUENCY_VECTOR))
BIN_MIN = FREQUENCY_VECTOR(LBound(FREQUENCY_VECTOR) + 1)
NBINS = FREQUENCY_VECTOR(LBound(FREQUENCY_VECTOR) + 2)

FREQUENCY_VECTOR = HISTOGRAM_FREQUENCY_FUNC(DATA_VECTOR, NBINS, BIN_MIN, BIN_WIDTH, 1)

ReDim SUMMARY_VECTOR(1 To 11, 1 To 2)
ReDim TEMP_MATRIX(0 To (NBINS + 1), 1 To 9)

TEMP_MATRIX(0, 1) = "LOWER LIMIT"
TEMP_MATRIX(0, 2) = "FREQ"
TEMP_MATRIX(0, 3) = "DISTR"
TEMP_MATRIX(0, 4) = "CUM DISTR"
TEMP_MATRIX(0, 5) = "1- CUM DISTR"
TEMP_MATRIX(0, 6) = "CUM CUM DISTR"
TEMP_MATRIX(0, 7) = "CUM 1 - CUM DISTR"
TEMP_MATRIX(0, 8) = "TARGET"
TEMP_MATRIX(0, 9) = "LOG OMEGA GAMMA"

TEMP_SUM = 0
For i = 1 To (NBINS + 1): TEMP_SUM = TEMP_SUM + FREQUENCY_VECTOR(i, 2): Next i

i = 1
TEMP_MATRIX(i, 1) = FREQUENCY_VECTOR(i, 1)
TEMP_MATRIX(i, 2) = FREQUENCY_VECTOR(i, 2)
TEMP_MATRIX(i, 3) = FREQUENCY_VECTOR(i, 2) / TEMP_SUM
TEMP_MATRIX(i, 4) = TEMP_MATRIX(i, 3)
TEMP_MATRIX(i, 5) = 1 - TEMP_MATRIX(i, 4)
TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 4)
TEMP_MATRIX(i, 7) = TEMP_MATRIX(i, 5)

For i = 2 To (NBINS + 1)
    TEMP_MATRIX(i, 1) = FREQUENCY_VECTOR(i, 1)
    TEMP_MATRIX(i, 2) = FREQUENCY_VECTOR(i, 2)
    TEMP_MATRIX(i, 3) = FREQUENCY_VECTOR(i, 2) / TEMP_SUM
    TEMP_MATRIX(i, 4) = TEMP_MATRIX(i, 3) + TEMP_MATRIX(i - 1, 4)
    TEMP_MATRIX(i, 5) = 1 - TEMP_MATRIX(i, 4)
    TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 4) + TEMP_MATRIX(i - 1, 6)
    TEMP_MATRIX(i, 7) = TEMP_MATRIX(i, 5) + TEMP_MATRIX(i - 1, 7)
Next i

For i = (NBINS + 1) To 1 Step -1
    If i < (NBINS + 1) Then
        If (TARGET_RATE >= TEMP_MATRIX(i, 1) And TARGET_RATE < TEMP_MATRIX(i + 1, 1)) Then
            TEMP_MATRIX(i, 8) = True
            j = i
        Else
            TEMP_MATRIX(i, 8) = False
        End If
    Else
        TEMP_MATRIX(i, 8) = False
    End If

    If (TEMP_MATRIX(i, 6) <> 0) And (IsNumeric(TEMP_MATRIX(i, 6))) Then
        TEMP_VAL = (TEMP_MATRIX((NBINS + 1), 7) - TEMP_MATRIX(i, 7)) / TEMP_MATRIX(i, 6)
        If TEMP_VAL > 0 Then
            TEMP_MATRIX(i, 9) = Log(TEMP_VAL)
        Else
            TEMP_MATRIX(i, 9) = ""
        End If
    Else
        TEMP_MATRIX(i, 9) = ""
    End If
Next i

If OUTPUT > 0 Then
    VECTOR_OMEGA_GAMMA_FUNC = TEMP_MATRIX
    Exit Function
End If

SUMMARY_VECTOR(6, 2) = 0
SUMMARY_VECTOR(7, 2) = 0
SUMMARY_VECTOR(8, 2) = 0
SUMMARY_VECTOR(9, 2) = 0

For i = 1 To NROWS
    If DATA_VECTOR(i, 1) >= TARGET_RATE Then: SUMMARY_VECTOR(6, 2) = SUMMARY_VECTOR(6, 2) + 1
    If DATA_VECTOR(i, 1) <= TARGET_RATE Then: SUMMARY_VECTOR(7, 2) = SUMMARY_VECTOR(7, 2) + 1
    If DATA_VECTOR(i, 1) >= TARGET_RATE Then: SUMMARY_VECTOR(8, 2) = SUMMARY_VECTOR(8, 2) + DATA_VECTOR(i, 1)
    If DATA_VECTOR(i, 1) <= TARGET_RATE Then: SUMMARY_VECTOR(9, 2) = SUMMARY_VECTOR(9, 2) + DATA_VECTOR(i, 1)
Next i

SUMMARY_VECTOR(1, 1) = "Gain Loss Ratio"
SUMMARY_VECTOR(1, 2) = SUMMARY_VECTOR(6, 2) / SUMMARY_VECTOR(7, 2)

SUMMARY_VECTOR(2, 1) = "Gains Losses"
SUMMARY_VECTOR(2, 2) = SUMMARY_VECTOR(8, 2) / SUMMARY_VECTOR(9, 2)

SUMMARY_VECTOR(3, 1) = "Expected Gains - Losses"
SUMMARY_VECTOR(3, 2) = ((SUMMARY_VECTOR(8, 2) / SUMMARY_VECTOR(6, 2)) / (SUMMARY_VECTOR(9, 2) / SUMMARY_VECTOR(7, 2)))

SUMMARY_VECTOR(4, 1) = "Gamma (Target)"
SUMMARY_VECTOR(4, 2) = ((SUMMARY_VECTOR(8, 2) / SUMMARY_VECTOR(6, 2)) / (SUMMARY_VECTOR(9, 2) / SUMMARY_VECTOR(7, 2))) * (1 - TEMP_MATRIX(j, 4)) / TEMP_MATRIX(j, 4)

SUMMARY_VECTOR(5, 1) = "OMEGA (Target)"
SUMMARY_VECTOR(5, 2) = Exp(TEMP_MATRIX(j, 9))

SUMMARY_VECTOR(6, 1) = "#Gains"
SUMMARY_VECTOR(7, 1) = "#Losses"
SUMMARY_VECTOR(8, 1) = "Summed Gains"
SUMMARY_VECTOR(9, 1) = "Summed Losses"

SUMMARY_VECTOR(10, 1) = "Expected Loss"
SUMMARY_VECTOR(10, 2) = SUMMARY_VECTOR(9, 2) / SUMMARY_VECTOR(7, 2)

SUMMARY_VECTOR(11, 1) = "Expected Gain"
SUMMARY_VECTOR(11, 2) = SUMMARY_VECTOR(8, 2) / SUMMARY_VECTOR(6, 2)

VECTOR_OMEGA_GAMMA_FUNC = SUMMARY_VECTOR

Exit Function
ERROR_LABEL:
VECTOR_OMEGA_GAMMA_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : VECTOR_OMEGA_EXPECTED_GL_FUNC
'DESCRIPTION   :

'Gains-to-Losses Measures
'There exists a large (and grwoing) number of statistics giving a feel
'for "what are the gains compared to the losses". For categorisation
'purposes, I prefer to group them under the header 'Gains-To-Losses Measures'.

'A lot of users (and authors) are confusing Gains-to-Losses measures with
'"Risk-Adjusted" measures. Only under very specific assumptions are
'gains-to-losses equivalent ro risk-adjusted performance.
'Gains-to-losses measures have the advantage that the actual paths of time
'series are discussed, and not just rather abstract characteristics of the
'return generating processes (like for example volatility). On the other
'hand, there are some rather esoteric measures in circulation, especially
'in the Hedge Fund industry.
 
'Gain-to-Loss Ratio
'GLR = count[G] / count[L]
 
'GLR... Gain:Loss Ratio
'count[.]... a counting function
'G... gains
'L... losses
'Of course, the beef is in the definition of "gains" and "losses". Example
'for gains are...
'"   Positive returns: G = count[r(p) >= 0]
'"   Excess returns over a benchmark (can be an index, a threshold etc.):
'G = count[r(p) >= r(b)]

'The disadvantage of the Gain:Loss Ratio is that the magnitude of the gains
'and losses are neglected: One can have three times as much gains than
'losses, if the magnitude of losses is ten times larger than the magnitude
'of gains, one probably not really want to invest.
 
'Gains-to-Losses
'The Gains:Losses statistic is basically a Gain:Loss Ratio taking into
'account the magnitude of the gains and losses...
'G:L = sum[G] / sum[L]
 
'G:L... Gains-to-Losses
'sum[.]... a summation function
'G... gains
'L... losses

'As said above, it is basically open how "gains" and "losses" are defined.
 
'Expected Gains-to-Losses
'A related measure is expected gains versus expected losses...
'eG:L = E[r|r>=L] / E[r|r<=L]  = avg[G] / avg[L]
 
'eG:L... Expected Gains:Losses
'E[.]... expectation operator (typically an arithmetic average)
'G... gains
'L... losses
 
'OMEGA
'Omega can be interpreted as some sort of 'probability-weighted ratio of gains over
'losses at a given level of expected return'.

'The risk-adjusted performance measures discussed before were all based on certain
'moments' (mean, variance etc.) of the return distribution of the portfolio or asset
'analyzed. This is only valid if the return distribution is fully defined by the
'moments used. If there exist features of the return distribution which are relevant
'to the risk and return preferences but not captured in the calculation of the
'measures, distortions are introduced.

'The so-called "Omega" measure tries to make use of the full return distribution and
'relies on very general assumptions about risk and return preferences: In order to
'be able to rank portfolios, the only rule necessary is that "more" is preferred to
'"less" ("non-satiation"). Further, Omega also takes into account  a level of return
'against which realized returns will be viewed as a "gain" or a "loss"; so in a
'certain way, Omega (and other measures discussed later) are sophisticated
'Gains-To-Losses statistics.

'Omega is most often encountered in the context of hedge funds, financial products
'which are designed and sold on the basis of their creative distribution and
'correlation characteristics.

'The way to Omega is through Gamma...

'G(L) = E[r|r>=L] / E[r|r<=L] *  {1 - F(L)} / F(L)
'r... fund returns
'L... threshold level
'G(L)... omega given a certain L
'F(r)... cumulative distribution function of returns, F(r) = Prob{r <= y}
'E[.]... expectation operator (usually an arithmetic average function)
 
'Omega is defined as...
'O(L) = int[1 - F(r)], L, b] / int[F(r)], a, L]
'r... fund returns
'L... threshold level
'O(L)... omega given a certain L
'F(r)... cummulative distribution function of returns, F(r) = Prob{r <= y}
'a... lower bound of the return distribution
'b... upper bound of the return distribution
'int[f, x, y]... integral of function f from x to y

'Omega is conceptually related to Stochastic Dominance (=using the full distribution)
'as well as to downside risk approaches (=taking into account a 'threshold').
'By construction, the value of Omega will always be one if the median of the
'distribution is chosen as the threshold return.

'Omega can be expressed has...
'O(l) = c(l) / p(l)
'C(L)... price of a european call with strike L
'P(L)... price of a european put with strike L

'LIBRARY       : STATISTICS
'GROUP         : OMEGA
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 14/02/2008
'************************************************************************************
'************************************************************************************

Function VECTOR_OMEGA_EXPECTED_GL_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal TARGET_RATE As Double = 1, _
Optional ByVal DATA_TYPE As Integer = 0, _
Optional ByVal LOG_SCALE As Integer = 0, _
Optional ByVal OUTPUT As Integer = 2)

Dim i As Long
Dim j As Long

Dim NROWS As Long

Dim NO_GAINS As Double
Dim NO_LOSSES As Double

Dim GAINS_VAL As Double
Dim LOSSES_VAL As Double

Dim EXPECTED_GAIN As Double
Dim EXPECTED_LOSS As Double

Dim GAIN_LOSS_RATIO As Double
Dim GAINS_LOSSES As Double
Dim EXPECTED_GAINS_LOSSES As Double

Dim TEMP_VECTOR As Variant
Dim TEMP_MATRIX As Variant
Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then: DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
If DATA_TYPE <> 0 Then: DATA_VECTOR = MATRIX_PERCENT_FUNC(DATA_VECTOR, LOG_SCALE)
DATA_VECTOR = MATRIX_QUICK_SORT_FUNC(DATA_VECTOR, 1, 1)

NROWS = UBound(DATA_VECTOR, 1)

ReDim TEMP_MATRIX(0 To NROWS, 1 To 10)

TEMP_MATRIX(0, 1) = "N"
TEMP_MATRIX(0, 2) = "Sorted Data"
TEMP_MATRIX(0, 3) = "F(x)"
TEMP_MATRIX(0, 4) = "1-F(x)"
TEMP_MATRIX(0, 5) = "Trapezium Rule F(x)"
TEMP_MATRIX(0, 6) = "Left Integral F(x)"
TEMP_MATRIX(0, 7) = "Trapezium Rule 1-F(x)"
TEMP_MATRIX(0, 8) = "Right Integral 1-F(x)"
TEMP_MATRIX(0, 9) = "Omega"
TEMP_MATRIX(0, 10) = "ln Omega"

TEMP_MATRIX(1, 1) = 1
TEMP_MATRIX(1, 2) = DATA_VECTOR(1, 1)
TEMP_MATRIX(1, 3) = 1 / NROWS
TEMP_MATRIX(1, 4) = 1 - TEMP_MATRIX(1, 3)
TEMP_MATRIX(1, 5) = TEMP_MATRIX(1, 3)
TEMP_MATRIX(1, 6) = TEMP_MATRIX(1, 5)
TEMP_MATRIX(1, 7) = TEMP_MATRIX(1, 4)

If TEMP_MATRIX(1, 2) > TARGET_RATE Then
    NO_GAINS = NO_GAINS + 1
    GAINS_VAL = GAINS_VAL + TEMP_MATRIX(1, 2)
Else
    NO_LOSSES = NO_LOSSES + 1
    LOSSES_VAL = LOSSES_VAL + TEMP_MATRIX(1, 2)
End If

For i = 2 To NROWS
    TEMP_MATRIX(i, 1) = i
    TEMP_MATRIX(i, 2) = DATA_VECTOR(i, 1)
    
    If TEMP_MATRIX(i, 2) > TARGET_RATE Then
        NO_GAINS = NO_GAINS + 1
        GAINS_VAL = GAINS_VAL + TEMP_MATRIX(i, 2)
    Else
        NO_LOSSES = NO_LOSSES + 1
        LOSSES_VAL = LOSSES_VAL + TEMP_MATRIX(i, 2)
    End If
    
    TEMP_MATRIX(i, 3) = i / NROWS
    TEMP_MATRIX(i, 4) = 1 - TEMP_MATRIX(i, 3)
    
    TEMP_MATRIX(i, 5) = 0.5 * (TEMP_MATRIX(i, 2) - TEMP_MATRIX(i - 1, 2)) * (TEMP_MATRIX(i, 3) + TEMP_MATRIX(i - 1, 3))
    TEMP_MATRIX(i, 6) = TEMP_MATRIX(i - 1, 6) + TEMP_MATRIX(i, 5)
    TEMP_MATRIX(i, 7) = 0.5 * (TEMP_MATRIX(i, 2) - TEMP_MATRIX(i - 1, 2)) * (TEMP_MATRIX(i, 4) + TEMP_MATRIX(i - 1, 4))
Next i

EXPECTED_GAIN = GAINS_VAL / NO_GAINS
EXPECTED_LOSS = LOSSES_VAL / NO_LOSSES
GAIN_LOSS_RATIO = NO_GAINS / NO_LOSSES
GAINS_LOSSES = GAINS_VAL / LOSSES_VAL
EXPECTED_GAINS_LOSSES = EXPECTED_GAIN / EXPECTED_LOSS

TEMP_MATRIX(NROWS, 8) = TEMP_MATRIX(NROWS, 7)
TEMP_MATRIX(NROWS, 9) = TEMP_MATRIX(NROWS, 8) / TEMP_MATRIX(NROWS, 6)
TEMP_MATRIX(NROWS, 10) = Log(TEMP_MATRIX(NROWS, 9))

For i = NROWS - 1 To 1 Step -1
    TEMP_MATRIX(i, 8) = TEMP_MATRIX(i + 1, 8) + TEMP_MATRIX(i, 7)
    TEMP_MATRIX(i, 9) = TEMP_MATRIX(i, 8) / TEMP_MATRIX(i, 6)
    TEMP_MATRIX(i, 10) = Log(TEMP_MATRIX(i, 9))
Next i

'-------------------------------------------------------------------------------
Select Case OUTPUT
'-------------------------------------------------------------------------------
Case 0
'-------------------------------------------------------------------------------
    VECTOR_OMEGA_EXPECTED_GL_FUNC = TEMP_MATRIX
'-------------------------------------------------------------------------------
Case 1
'-------------------------------------------------------------------------------
    ReDim TEMP_VECTOR(1 To 9, 1 To 3)
    
    TEMP_VECTOR(1, 1) = "--"
    TEMP_VECTOR(1, 2) = "Omega Interpolation"
    TEMP_VECTOR(1, 3) = "F(x)Interpolation"
    
    TEMP_VECTOR(2, 1) = "x"
    TEMP_VECTOR(2, 2) = TARGET_RATE
    TEMP_VECTOR(2, 3) = ""
    
    TEMP_VECTOR(3, 1) = "x0"
    TEMP_VECTOR(3, 2) = ""
    TEMP_VECTOR(3, 3) = ""
    
    TEMP_VECTOR(4, 1) = "x1"
    TEMP_VECTOR(4, 2) = ""
    TEMP_VECTOR(4, 3) = ""
    
    TEMP_VECTOR(5, 1) = "y0"
    TEMP_VECTOR(5, 2) = ""
    TEMP_VECTOR(5, 3) = ""
    
    TEMP_VECTOR(6, 1) = "y1"
    TEMP_VECTOR(6, 2) = ""
    TEMP_VECTOR(6, 3) = ""
    
    j = 1
    For i = 1 To NROWS - 1
        If (TARGET_RATE >= TEMP_MATRIX(i, 2) And TARGET_RATE < TEMP_MATRIX(i + 1, 2)) Then: j = i
    Next i
    TEMP_VECTOR(3, 2) = TEMP_MATRIX(j, 2)
    TEMP_VECTOR(4, 2) = TEMP_MATRIX(j + 1, 2)
    
    TEMP_VECTOR(5, 2) = TEMP_MATRIX(j, 9)
    TEMP_VECTOR(5, 3) = TEMP_MATRIX(j, 3)
    
    TEMP_VECTOR(6, 2) = TEMP_MATRIX(j + 1, 9)
    TEMP_VECTOR(6, 3) = TEMP_MATRIX(j + 1, 3)
    
    TEMP_VECTOR(7, 1) = "A = (X - x0) / (x1 - x0)"
    TEMP_VECTOR(7, 2) = (TEMP_VECTOR(2, 2) - TEMP_VECTOR(3, 2)) / (TEMP_VECTOR(4, 2) - TEMP_VECTOR(3, 2))
    TEMP_VECTOR(7, 3) = (TEMP_VECTOR(2, 2) - TEMP_VECTOR(3, 2)) / (TEMP_VECTOR(4, 2) - TEMP_VECTOR(3, 2))
    
    TEMP_VECTOR(8, 1) = "Omega {y = y0 + A * (y1 - y0)}"
    TEMP_VECTOR(8, 2) = TEMP_VECTOR(5, 2) + TEMP_VECTOR(7, 3) * (TEMP_VECTOR(6, 2) - TEMP_VECTOR(5, 2))
    TEMP_VECTOR(8, 3) = TEMP_VECTOR(5, 3) + TEMP_VECTOR(7, 3) * (TEMP_VECTOR(6, 3) - TEMP_VECTOR(5, 3))
    
    TEMP_VECTOR(9, 1) = "Gamma"
    TEMP_VECTOR(9, 2) = EXPECTED_GAINS_LOSSES * (1 - TEMP_VECTOR(8, 3)) / TEMP_VECTOR(8, 3)
    TEMP_VECTOR(9, 3) = ""
    
    VECTOR_OMEGA_EXPECTED_GL_FUNC = TEMP_VECTOR
'-------------------------------------------------------------------------------
Case Else
'-------------------------------------------------------------------------------
    ReDim TEMP_VECTOR(1 To 9, 1 To 2)
    TEMP_VECTOR(1, 1) = "#Gains"
    TEMP_VECTOR(1, 2) = NO_GAINS

    TEMP_VECTOR(2, 1) = "#Losses"
    TEMP_VECTOR(2, 2) = NO_LOSSES

    TEMP_VECTOR(3, 1) = "Summed Gains"
    TEMP_VECTOR(3, 2) = GAINS_VAL

    TEMP_VECTOR(4, 1) = "Summed Losses"
    TEMP_VECTOR(4, 2) = LOSSES_VAL

    TEMP_VECTOR(5, 1) = "Expected Gain"
    TEMP_VECTOR(5, 2) = EXPECTED_GAIN

    TEMP_VECTOR(6, 1) = "Expected Loss"
    TEMP_VECTOR(6, 2) = EXPECTED_LOSS

    TEMP_VECTOR(7, 1) = "Gain - Loss Ratio"
    TEMP_VECTOR(7, 2) = GAIN_LOSS_RATIO

    TEMP_VECTOR(8, 1) = "Gains - Losses"
    TEMP_VECTOR(8, 2) = GAINS_LOSSES

    TEMP_VECTOR(9, 1) = "Expected Gains - Losses"
    TEMP_VECTOR(9, 2) = EXPECTED_GAINS_LOSSES

    VECTOR_OMEGA_EXPECTED_GL_FUNC = TEMP_VECTOR
'-------------------------------------------------------------------------------
End Select
'-------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
VECTOR_OMEGA_EXPECTED_GL_FUNC = Err.number
End Function
