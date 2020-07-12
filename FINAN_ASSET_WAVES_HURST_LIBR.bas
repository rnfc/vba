Attribute VB_Name = "FINAN_ASSET_WAVES_HURST_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'-------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------
'Once upon a time, a British government bureaucrat named Harold Edwin Hurst studied 800 years
'of records of the Nile's flooding. He noticed that there was a tendency for a high flood year
'to be followed by another high flood year, and for a low flood year to be followed by another
'low flood year. Was that accidental ... or was there really some correlation between levels?
'Did the height at year 5 have an effect on the height in year 6?
'Are you talking about river levels ... or financial stuff?

'To analyze, we might do something like this:
'   1. Note the heights of the n flood levels:
'            h(1), h(2), ... h(n)
'   2. Let m be the Mean of these levels:
'            M = (1/n) [ h(1)+h(2)+...+h(n) ]
'   3. Calculate the deviations from the mean:
'            x(1) = h(1) - m
'            x(2) = h(2) - m
'            ...
'            x(n) = h(n) - m
'      Note that the set of xs have zero mean.
'      Positive x 's indicate that the Nile level was above the average.
'   4. Now calculate the Sums:
'            Y(1) = x(1)
'            Y(2) = x(1) + x(2)
'            ...
'            Y(n) = x(1) + x(2) + ...+ x(n)
'      Note that the set of partial sums, the Y's, are sums of zero-mean variables.
'      They will be positive if there's a preponderance of positive x's.
'      Note, too, that Y(k) = Y(k-1) + x(k).
'   5. Let R(n) = MAX[Y(k)] - MIN[Y(k)]
'      This difference between the maximum and minimum of the n values is called the Range
'   6. Let s(n) be the standard deviation of the set of n h-values.

'As it turns out, the probability theorist William Feller proved that if a series of random
'variables (like the x's) had finite standard deviation and were independent, then the
'so-called R/s statistic (formed over n observations) would increase in proportion to
'n1/2 (for large values of n). This guy, R/s, is called the rescaled range
'Anyway, we now have:
'      R(n) / s(n) - kn^1/2    ... where k is some constant
'If that were true, then we'd expect that:
'      log(R/s ) ? log(k) + (1/2) log(n)
'So, if we were to plot log(R/s ) vs log(n), we'd expect it to be approximately a straight
'line with slope (1/2).

'Anyway, what Hurst apparently found, was that the plot had a slope closer to 0.7 (rather than 0.5).
'So, what's that mean?
'I guess it means that the annual Nile levels weren't independent, but this year's level might be
'expected to affect next year's level. Indeed, if the slope of the log(R / s ) vs log(n)
'"best fit line" is H, then we'd expect: R / s - kn^H

'The interesting thing is that many things seem to exhibit this long term patterns or dependence ...
'seven years of plenty followed by seven years of plenty.
'-------------------------------------------------------------------------------------------------
'References:
'-------------------------------------------------------------------------------------------------
'http://www.gummy-stuff.org/hurst.htm
'http://www.esr.ie/vol31_3/3Conniffe.pdf
'http://www-history.mcs.st-and.ac.uk/history/Mathematicians/Feller.html
'http://www.bearcave.com/misl/misl_tech/wavelets/hurst/
'http://www.gummy-stuff.org/coin-tossing.htm
'http://www.gummy-stuff.org/RTM.htm
'-------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------

Function ASSET_HURST_EXPONENT_FUNC(ByVal TICKER_STR As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal END_DATE As Date, _
Optional ByVal S_VAL As Long = 500, _
Optional ByVal D_VAL As Long = 5, _
Optional ByVal M_VAL As Long = 100, _
Optional ByVal OUTPUT As Integer = 0)

'S_VAL --> Start at Day
'D_VAL --> increment

Dim i As Long
Dim j As Long
Dim k As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim R_VAL As Double
Dim RS_VAL As Double 'R/s

Dim YMIN_VAL As Double
Dim YMAX_VAL As Double
Dim TEMP_SUM As Double

Dim MEAN_VAL As Double
Dim SIGMA_VAL As Double

Dim SLOPE_VAL As Double
Dim INTERCEPT_VAL As Double

Dim XDATA_VECTOR As Variant
Dim YDATA_VECTOR As Variant

Dim DATA_MATRIX As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

If IsArray(TICKER_STR) = False Then
    DATA_MATRIX = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, START_DATE, END_DATE, _
                  "d", "DOHLCVA", False, True, True)
Else
    DATA_MATRIX = TICKER_STR
End If

NROWS = UBound(DATA_MATRIX, 1)
If S_VAL + D_VAL * M_VAL > NROWS Then: GoTo ERROR_LABEL 'too large
NCOLUMNS = UBound(DATA_MATRIX, 2)

ReDim TEMP_MATRIX(0 To NROWS, 1 To 15)
i = 0
TEMP_MATRIX(i, 1) = "DATE"
TEMP_MATRIX(i, 2) = "OPEN"
TEMP_MATRIX(i, 3) = "HIGH"
TEMP_MATRIX(i, 4) = "LOW"
TEMP_MATRIX(i, 5) = "CLOSE"
TEMP_MATRIX(i, 6) = "VOLUME"
TEMP_MATRIX(i, 7) = "ADJ_CLOSE"
TEMP_MATRIX(i, 8) = "RETURN"
TEMP_MATRIX(i, 9) = "R(k)"
TEMP_MATRIX(i, 10) = "X(k)"
TEMP_MATRIX(i, 11) = "Y(k)"
TEMP_MATRIX(i, 12) = "N"
TEMP_MATRIX(i, 13) = "log(N)"
TEMP_MATRIX(i, 14) = "log(R/S)"
TEMP_MATRIX(i, 15) = "R/S = "

i = 1
For j = 1 To NCOLUMNS: TEMP_MATRIX(i, j) = DATA_MATRIX(i, j): Next j
TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 6) / 1000
TEMP_MATRIX(i, 8) = TEMP_MATRIX(i, 5) / TEMP_MATRIX(i, 2) - 1

For i = 2 To NROWS
    For j = 1 To NCOLUMNS: TEMP_MATRIX(i, j) = DATA_MATRIX(i, j): Next j
    TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 6) / 1000
    TEMP_MATRIX(i, 8) = TEMP_MATRIX(i, 7) / TEMP_MATRIX(i - 1, 7) - 1
Next i

If OUTPUT > 0 Then
    ReDim XDATA_VECTOR(1 To M_VAL, 1 To 1)
    ReDim YDATA_VECTOR(1 To M_VAL, 1 To 1)
End If

j = S_VAL
For k = 1 To M_VAL
    j = j + D_VAL
    TEMP_SUM = 0
    For i = 1 To j
        TEMP_MATRIX(i, 9) = TEMP_MATRIX(i, 8)
        TEMP_SUM = TEMP_SUM + TEMP_MATRIX(i, 9)
    Next i
    MEAN_VAL = TEMP_SUM / j
    TEMP_SUM = 0
    For i = 1 To j: TEMP_SUM = TEMP_SUM + (TEMP_MATRIX(i, 9) - MEAN_VAL) ^ 2: Next i
    SIGMA_VAL = (TEMP_SUM / j) ^ 0.5
    i = 1
    TEMP_MATRIX(i, 10) = TEMP_MATRIX(i, 9) - MEAN_VAL
    TEMP_MATRIX(i, 11) = TEMP_MATRIX(i, 10)
    YMIN_VAL = TEMP_MATRIX(i, 11): YMAX_VAL = TEMP_MATRIX(i, 11)
    For i = 2 To j
        TEMP_MATRIX(i, 10) = TEMP_MATRIX(i, 9) - MEAN_VAL
        TEMP_MATRIX(i, 11) = TEMP_MATRIX(i - 1, 11) + TEMP_MATRIX(i, 10)
        If TEMP_MATRIX(i, 11) < YMIN_VAL Then: YMIN_VAL = TEMP_MATRIX(i, 11)
        If TEMP_MATRIX(i, 11) > YMAX_VAL Then: YMAX_VAL = TEMP_MATRIX(i, 11)
    Next i
    R_VAL = YMAX_VAL - YMIN_VAL
    RS_VAL = R_VAL / SIGMA_VAL
    
    TEMP_MATRIX(k, 12) = TEMP_MATRIX(j, 1) 'j
    TEMP_MATRIX(k, 13) = Log(j)
    TEMP_MATRIX(k, 15) = RS_VAL
    TEMP_MATRIX(k, 14) = Log(TEMP_MATRIX(k, 15))
    
    If OUTPUT > 0 Then
        XDATA_VECTOR(k, 1) = TEMP_MATRIX(k, 13)
        YDATA_VECTOR(k, 1) = TEMP_MATRIX(k, 14)
    End If
Next k
For i = M_VAL + 1 To NROWS: For j = 12 To 15: TEMP_MATRIX(i, j) = "": Next j: Next i
i = 0
TEMP_MATRIX(i, 15) = TEMP_MATRIX(i, 15) & Format(RS_VAL, "0.0000")

'-----------------------------------------------------------------------------------
Select Case OUTPUT
'-----------------------------------------------------------------------------------
Case 0
'-----------------------------------------------------------------------------------
    ASSET_HURST_EXPONENT_FUNC = TEMP_MATRIX
'-----------------------------------------------------------------------------------
Case Else
    DATA_MATRIX = REGRESSION_SIMPLE_COEF_FUNC(XDATA_VECTOR, YDATA_VECTOR)
    SLOPE_VAL = DATA_MATRIX(1, 1) 'Hurst Exponent
    INTERCEPT_VAL = DATA_MATRIX(2, 1)
    ASSET_HURST_EXPONENT_FUNC = Array(RS_VAL, SLOPE_VAL, INTERCEPT_VAL)
'-----------------------------------------------------------------------------------
End Select
'-----------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
ASSET_HURST_EXPONENT_FUNC = Err.number
End Function
