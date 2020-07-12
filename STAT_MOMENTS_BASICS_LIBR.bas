Attribute VB_Name = "STAT_MOMENTS_BA_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : DATA_BASIC_MOMENTS_FRAME_FUNC
'DESCRIPTION   : RETURNS A MSGBOX WITH THE DESCRIPTIVE STATISTICS OF A DATA-SET
'LIBRARY       : STATISTICS
'GROUP         : MOMENTS
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 14/02/2008
'************************************************************************************
'************************************************************************************

Function DATA_BASIC_MOMENTS_FRAME_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal DATA_TYPE As Integer = 0, _
Optional ByVal LOG_SCALE As Integer = 0, _
Optional ByVal CI_FACTOR As Double = 0.05, _
Optional ByVal SCOLUMN As Long = 1)

Dim j As Long
Dim NCOLUMNS As Long

Dim MSG_STR As String
Dim TEMP_STYLE As Variant
Dim TEMP_TITLE As Variant
'Dim HELP As Variant
'Dim CTXT As Variant
Dim TEMP_FLAG As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_BASIC_MOMENTS_FRAME_FUNC = False

DATA_MATRIX = DATA_RNG
NCOLUMNS = UBound(DATA_MATRIX, 2)

DATA_MATRIX = DATA_BASIC_MOMENTS_FUNC(DATA_MATRIX, DATA_TYPE, LOG_SCALE, CI_FACTOR, 0)

For j = SCOLUMN To NCOLUMNS
    MSG_STR = "OBS:" & vbTab & Format(DATA_MATRIX(j, 1), "#,0.000") & vbCrLf & _
              "MIN:" & vbTab & Format(DATA_MATRIX(j, 2), "#,0.000") & vbCrLf & _
              "MAX:" & vbTab & Format(DATA_MATRIX(j, 3), "#,0.000") & vbCrLf & _
              "MEAN:" & vbTab & Format(DATA_MATRIX(j, 4), "#,0.000") & vbCrLf & _
              "ADEV:" & vbTab & Format(DATA_MATRIX(j, 5), "#,0.000") & vbCrLf & _
              "VAR:" & vbTab & Format(DATA_MATRIX(j, 6), "#,0.000") & vbCrLf & _
              "STDEV:" & vbTab & Format(DATA_MATRIX(j, 7), "#,0.000") & vbCrLf & _
              "STDEVP:" & vbTab & Format(DATA_MATRIX(j, 8), "#,0.000") & vbCrLf & _
              "SE:" & vbTab & Format(DATA_MATRIX(j, 9), "#,0.000") & vbCrLf & _
              "UPPER CI:" & vbTab & Format(DATA_MATRIX(j, 10), "#,0.000") & vbCrLf & _
              "LOWER CI:" & vbTab & Format(DATA_MATRIX(j, 11), "#,0.000") & vbCrLf & _
              "SKEW:" & vbTab & Format(DATA_MATRIX(j, 12), "#,0.000") & vbCrLf & _
              "KURT:" & vbTab & Format(DATA_MATRIX(j, 13), "#,0.000") & vbCrLf

    TEMP_STYLE = vbOKCancel + vbInformation + vbDefaultButton2    ' Define buttons.
    TEMP_TITLE = "DESCRIPTIVE STAT"    ' Define title.
    'HELP = "DEMO.HLP"    ' Define Help file.
    'CTXT = 1000    ' Define topic
    TEMP_FLAG = MsgBox(MSG_STR, TEMP_STYLE, TEMP_TITLE) 'HELP, CTXT)
    If TEMP_FLAG = vbCancel Then: Exit For
Next j

DATA_BASIC_MOMENTS_FRAME_FUNC = True

Exit Function
ERROR_LABEL:
DATA_BASIC_MOMENTS_FRAME_FUNC = False
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : DATA_BASIC_MOMENTS_FUNC
'DESCRIPTION   : RETURNS THE DESCRIPTIVE STATISTICS OF A DATA-SET
'LIBRARY       : STATISTICS
'GROUP         : MOMENTS
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 14/02/2008
'************************************************************************************
'************************************************************************************

Function DATA_BASIC_MOMENTS_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal DATA_TYPE As Integer = 0, _
Optional ByVal LOG_SCALE As Integer = 0, _
Optional ByVal CI_FACTOR As Double = 0.05, _
Optional ByVal OUTPUT As Integer = 0)

'For Simulations
    'NCOLUMNS = DATES
    'NROWS = REPETITION PER DAY

'For Historical DATA SETS
    'NCOLUMNS = ASSETS
    'NROWS = Returns per day
    
Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim MIN_VAL As Double
Dim MAX_VAL As Double

Dim DEV_VAL As Double
Dim VAR_VAL As Double
Dim SKEW_VAL As Double
Dim KURT_VAL As Double

Dim TEMP_VAL As Double
Dim TEMP1_SUM As Double
Dim TEMP2_SUM As Double

Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

Const epsilon As Double = 10 ^ -15

On Error GoTo ERROR_LABEL

'---------------------------------------------------------------------------------
'Volatility is by far the most often used absolute risk measure in
'financial applications. Besides conceptual issues (do people really
'perceive positive deviations from the mean as "risk"?), the major
'disadvantage of the volatility measure is lack of robustness: By
'calculating the squared deviation from the mean, outliers receive
'much more weight than realizsations closer to the mean.

'An robust alternative to volatility is the so-called interquartile range...

'InterR = 3rd Quartile - 1st Quartile

'3rd Quartile... Value at which 75% of all observations are to the right
'and 25% to the right.

'3rd Quartile... Value at which 25% of all observations are to the right
'and 75% to the right.
'---------------------------------------------------------------------------------

DATA_MATRIX = DATA_RNG
If DATA_TYPE <> 0 Then DATA_MATRIX = MATRIX_PERCENT_FUNC(DATA_MATRIX, LOG_SCALE)

NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

If CI_FACTOR >= 1 Then: CI_FACTOR = 1 - epsilon '0.999999999999999
If CI_FACTOR <= 0 Then: CI_FACTOR = epsilon '0.000000000000001

TEMP_VAL = NORMSINV_FUNC(1 - CI_FACTOR / 2, 0, 1, 0)

ReDim TEMP_MATRIX(1 To NCOLUMNS, 1 To 13)

For j = 1 To NCOLUMNS
    MIN_VAL = 0: MAX_VAL = 0
    TEMP1_SUM = 0
    
    MAX_VAL = DATA_MATRIX(1, j): MIN_VAL = DATA_MATRIX(1, j)
    For i = 1 To NROWS
        TEMP1_SUM = TEMP1_SUM + DATA_MATRIX(i, j)
        If MAX_VAL < DATA_MATRIX(i, j) Then MAX_VAL = DATA_MATRIX(i, j)
        If MIN_VAL > DATA_MATRIX(i, j) Then MIN_VAL = DATA_MATRIX(i, j)
    Next i
    TEMP_MATRIX(j, 1) = NROWS 'OBSERVATIONS
    TEMP_MATRIX(j, 2) = MIN_VAL ' MIN
    TEMP_MATRIX(j, 3) = MAX_VAL 'MAX
    TEMP_MATRIX(j, 4) = TEMP1_SUM / NROWS 'MEAN
        
    DEV_VAL = 0: TEMP2_SUM = 0
    VAR_VAL = 0: TEMP1_SUM = 0
    For i = 1 To NROWS
        DEV_VAL = (DATA_MATRIX(i, j) - TEMP_MATRIX(j, 4))
        TEMP1_SUM = TEMP1_SUM + DEV_VAL

        TEMP2_SUM = TEMP2_SUM + Abs(DEV_VAL)
        VAR_VAL = DEV_VAL * DEV_VAL + VAR_VAL
    Next i
        
    TEMP_MATRIX(j, 5) = TEMP2_SUM / NROWS 'AVERAGE DEVIATION
    TEMP_MATRIX(j, 6) = (VAR_VAL - TEMP1_SUM * TEMP1_SUM / NROWS) / (NROWS - 1) 'Variance: Corrected two-pass formula.

    TEMP_MATRIX(j, 7) = Sqr(TEMP_MATRIX(j, 6)) 'Sample Standard Deviation
    TEMP_MATRIX(j, 8) = Sqr(NROWS / (NROWS - 1)) * TEMP_MATRIX(j, 7) 'Population Standard Deviation
    'TEMP_MATRIX(j, 8) = Sqr((VAR_VAL / NROWS)) 'Population Standard Deviation - According to Excel.
    TEMP_MATRIX(j, 9) = TEMP_MATRIX(j, 7) / Sqr(NROWS) 'Standard Error
    TEMP_MATRIX(j, 10) = TEMP_MATRIX(j, 4) + TEMP_VAL * TEMP_MATRIX(j, 9) 'Upper Confidence Interval
    TEMP_MATRIX(j, 11) = TEMP_MATRIX(j, 4) - TEMP_VAL * TEMP_MATRIX(j, 9) 'Lower Confidence Interval
'--------------Calculate 3rd and 4th moments (skewness and kurtosis)
    If (TEMP_MATRIX(j, 6) <> 0) Then
        On Error GoTo 0
        On Error GoTo 1983: 'Fixing Error Traps from Overflows
        SKEW_VAL = 0: KURT_VAL = 0
        For i = 1 To NROWS
            SKEW_VAL = SKEW_VAL + ((DATA_MATRIX(i, j) - TEMP_MATRIX(j, 4)) / TEMP_MATRIX(j, 7)) ^ 3
            KURT_VAL = KURT_VAL + ((DATA_MATRIX(i, j) - TEMP_MATRIX(j, 4)) / TEMP_MATRIX(j, 7)) ^ 4
        Next i
        TEMP_MATRIX(j, 12) = SKEW_VAL / NROWS
        TEMP_MATRIX(j, 13) = KURT_VAL / NROWS - 3
'        TEMP_MATRIX(j, 12) = SKEW_VAL * (NROWS / ((NROWS - 1) * (NROWS - 2))) 'Excel Definition
'        TEMP_MATRIX(j, 13) = (KURT_VAL * (NROWS * (NROWS + 1) / ((NROWS - 1) * (NROWS - 2) * (NROWS - 3)))) - ((3 * (NROWS - 1) ^ 2 / ((NROWS - 2) * (NROWS - 3)))) 'Excel Definition
    Else
1983:   TEMP_MATRIX(j, 12) = 0 'SKEW
        TEMP_MATRIX(j, 13) = 0 'KURTOSIS
    End If
Next j

Select Case OUTPUT
Case 0
    DATA_BASIC_MOMENTS_FUNC = TEMP_MATRIX
Case Else
    TEMP_MATRIX = MATRIX_ADD_ROWS_FUNC(TEMP_MATRIX, 1, 1)
    TEMP_MATRIX(1, 1) = "OBSERVATIONS"
    TEMP_MATRIX(1, 2) = "MIN"
    TEMP_MATRIX(1, 3) = "MAX"
    TEMP_MATRIX(1, 4) = "MEAN"
    TEMP_MATRIX(1, 5) = "AVERAGE DEVIATION"
    TEMP_MATRIX(1, 6) = "VARIANCE"
    TEMP_MATRIX(1, 7) = "SAMPLE STANDARD DEVIATION"
    TEMP_MATRIX(1, 8) = "STANDARD DEVIATION POPULATION"
    TEMP_MATRIX(1, 9) = "STANDARD ERROR"
    TEMP_MATRIX(1, 10) = "UPPER CONFIDENCE INTERVAL"
    TEMP_MATRIX(1, 11) = "LOWER CONFIDENCE INTERVAL"
    TEMP_MATRIX(1, 12) = "SKEW"
    TEMP_MATRIX(1, 13) = "KURT"
    DATA_BASIC_MOMENTS_FUNC = TEMP_MATRIX
End Select

Exit Function
ERROR_LABEL:
DATA_BASIC_MOMENTS_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : DATA_ADVANCED_MOMENTS_FUNC
'DESCRIPTION   : RETURNS THE ADVANCED DESCRIPTIVE STATISTICS OF A DATA-SET
'LIBRARY       : STATISTICS
'GROUP         : MOMENTS
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 14/02/2008
'************************************************************************************
'************************************************************************************

Function DATA_ADVANCED_MOMENTS_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal DATA_TYPE As Integer = 0, _
Optional ByVal LOG_SCALE As Integer = 0, _
Optional ByVal CI_FACTOR As Double = 0.05, _
Optional ByVal OUTPUT As Integer = 0)

'LAST OBSERVATIONS OF MSCI:

'1) A large number of hedge fund indices exhibit low correlations with
'   MSCI World and statistically non-normal returns.

'2) A small number of hedge fund indices exhibit significant cokewness
'   and cokurtosis with MSCI World

'3) 100% of the indices exhibiting significant coskewness have negative
'   coskewness, depsite the fact that customer can be expected to have
'   preferences for positive cokurtosis

'4) 30% of the indices exhibiting significant cokurtosis have negative
'   cokurtosis(=fat left tails). Customers can be expected to have a
'   preference for positive cokurtosis (=fat right tails=positive
'   excess returns)

Dim j As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_ARR As Variant
Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

Dim TEMP1_VECTOR As Variant
Dim TEMP2_VECTOR As Variant

On Error Resume Next 'On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
If UBound(DATA_MATRIX, 1) = 1 Then: DATA_MATRIX = MATRIX_TRANSPOSE_FUNC(DATA_MATRIX)
If DATA_TYPE <> 0 Then DATA_MATRIX = MATRIX_PERCENT_FUNC(DATA_MATRIX, LOG_SCALE)

NCOLUMNS = UBound(DATA_MATRIX, 2)
NROWS = UBound(DATA_MATRIX, 1)


ReDim TEMP_MATRIX(1 To 15, 1 To NCOLUMNS + 1)

TEMP_MATRIX(1, 1) = "SKEWNESS"
TEMP_MATRIX(2, 1) = "BOWLEY SKEW"
TEMP_MATRIX(3, 1) = "PEARSON SKEW"
TEMP_MATRIX(4, 1) = "KURTOSIS"

TEMP_MATRIX(5, 1) = "MOORE KURTOSIS"
TEMP_MATRIX(6, 1) = "CROW/SIDDIQUI KURTOSIS"
TEMP_MATRIX(7, 1) = "PEAKEDNESS"
TEMP_MATRIX(8, 1) = "TAILS"

TEMP_MATRIX(9, 1) = "L = PEAK * TAIL"
TEMP_MATRIX(10, 1) = "PPC^2"
TEMP_MATRIX(11, 1) = "JB P - VALUE"
TEMP_MATRIX(12, 1) = "STANDARD DEVIATION"

TEMP_MATRIX(13, 1) = "INTERQUARTILE RANGE"
TEMP_MATRIX(14, 1) = "AVERAGE"
TEMP_MATRIX(15, 1) = "MEDIAN"

For j = 1 To NCOLUMNS
    
    TEMP1_VECTOR = MATRIX_GET_COLUMN_FUNC(DATA_MATRIX, j, 1)
    TEMP1_VECTOR = VECTOR_TRIM_FUNC(TEMP1_VECTOR, "")
    TEMP2_VECTOR = DATA_BASIC_MOMENTS_FUNC(TEMP1_VECTOR, 0, 0, CI_FACTOR, 0)
    TEMP_MATRIX(1, j + 1) = TEMP2_VECTOR(1, 12)
    TEMP_ARR = SKEW_BOWLEY_FUNC(TEMP1_VECTOR, 2)
    TEMP_MATRIX(2, j + 1) = TEMP_ARR(LBound(TEMP_ARR))
    TEMP_MATRIX(3, j + 1) = SKEW_PEARSON_FUNC(TEMP1_VECTOR)
    TEMP_MATRIX(4, j + 1) = TEMP2_VECTOR(1, 13)
    TEMP_MATRIX(5, j + 1) = KURT_MOORE_FUNC(TEMP1_VECTOR)
    TEMP_MATRIX(6, j + 1) = KURT_CROW_FUNC(TEMP1_VECTOR)
    TEMP_MATRIX(7, j + 1) = PEAKEDNESS_FUNC(TEMP1_VECTOR, 0.125, 0.25)
    TEMP_MATRIX(8, j + 1) = TAIL_WEIGHT_FUNC(TEMP1_VECTOR, 0.025, 0.125)
    TEMP_MATRIX(9, j + 1) = PEAK_TAIL_FUNC(TEMP1_VECTOR, 0.025, 0.125, 0.25)
    TEMP_MATRIX(10, j + 1) = CORRELATION_PROBABILITY_PLOT_FUNC(TEMP1_VECTOR) ^ 2
    TEMP_MATRIX(11, j + 1) = JARQUE_BERA_HYPOTHESIS_MODEL_FUNC(TEMP2_VECTOR(1, 12), TEMP2_VECTOR(1, 13), TEMP2_VECTOR(1, 1), CI_FACTOR, 2)
    TEMP_MATRIX(12, j + 1) = TEMP2_VECTOR(1, 7)
    TEMP_MATRIX(13, j + 1) = TEMP_ARR(UBound(TEMP_ARR))
    TEMP_MATRIX(14, j + 1) = TEMP2_VECTOR(1, 4)
    TEMP_MATRIX(15, j + 1) = HISTOGRAM_PERCENTILE_FUNC(TEMP1_VECTOR, 0.5, 1)
Next j

Select Case OUTPUT
Case 0
    DATA_ADVANCED_MOMENTS_FUNC = MATRIX_TRANSPOSE_FUNC(TEMP_MATRIX)
Case Else
    DATA_ADVANCED_MOMENTS_FUNC = TEMP_MATRIX
End Select

Exit Function
ERROR_LABEL:
DATA_ADVANCED_MOMENTS_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : FOUR_MOMENTS_INDEX_FUNC

'DESCRIPTION   : When financial assets are normally distributed, the historical
'asset return, the asset standard deviation and its covariance with the
'market are enough to estimate the asset expected return. This model is
'the 2-Moment CAPM developed by Sharpe (1964), Lintner (1965) and Mossin
'(1966). We claim that the risk in not only in volatility and linear
'correlation, but in skewness, kurtosis, systematic skewness, and systematic
'kurtosis. The model developed below and applied in HFOptimizer platform is
'the Four Moment CAPM, which accounts for the volatility, the skewness, the
'kurtosis of the assets, and of the world equity market and their linear
'(BETA) or no-linear (coskewness, cokurtosis) dependencies.

'REFERENCE: This Four-Moment Capital Asset Pricing Function is based on
'two academic papers recently issued (Jurcenzko and Maillet, The Four
'Moment CAPM: Some Basic Results , working paper, 2002, and Hwang and
'Satchell, Modeling Emerging Market Risk Premia Using Higher Moments ,
'working paper, 1999).

'LIBRARY       : STATISTICS
'GROUP         : MOMENTS
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 14/02/2008
'************************************************************************************
'************************************************************************************

Function FOUR_MOMENTS_INDEX_FUNC(ByRef DATA_RNG As Variant, _
ByRef XDATA_RNG As Variant, _
Optional ByVal DATA_TYPE As Integer = 0, _
Optional ByVal LOG_SCALE As Integer = 0, _
Optional ByVal CI_FACTOR As Double = 0.05, _
Optional ByVal OUTPUT As Integer = 1)

Dim j As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_VAL As Double
Dim TEMP_GROUP As Variant
Dim TEMP_MATRIX As Variant
Dim TEMP1_VECTOR As Variant
Dim TEMP2_VECTOR As Variant

Dim XDATA_VECTOR As Variant
Dim DATA_MATRIX As Variant

On Error Resume Next 'On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
If UBound(DATA_MATRIX, 1) = 1 Then: DATA_MATRIX = MATRIX_TRANSPOSE_FUNC(DATA_MATRIX)

XDATA_VECTOR = XDATA_RNG
If UBound(XDATA_VECTOR, 1) = 1 Then: XDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(XDATA_VECTOR)

If UBound(DATA_MATRIX, 1) <> UBound(XDATA_VECTOR, 1) Then: GoTo ERROR_LABEL

If DATA_TYPE <> 0 Then DATA_MATRIX = MATRIX_PERCENT_FUNC(DATA_MATRIX, LOG_SCALE)
If DATA_TYPE <> 0 Then XDATA_VECTOR = MATRIX_PERCENT_FUNC(XDATA_VECTOR, LOG_SCALE)

NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

ReDim TEMP_MATRIX(1 To 19, 1 To NCOLUMNS + 1)

TEMP_MATRIX(1, 1) = "Skewness"
TEMP_MATRIX(2, 1) = "Kurtosis"
TEMP_MATRIX(3, 1) = "Is Jarque-Bera Normal?"
TEMP_MATRIX(4, 1) = "Jarque-Bera p-Value"

TEMP_MATRIX(5, 1) = "R^2"
TEMP_MATRIX(6, 1) = "a1: Alpha"
TEMP_MATRIX(7, 1) = "a2: Beta"
TEMP_MATRIX(8, 1) = "a3: Coskewness"
TEMP_MATRIX(9, 1) = "a4: Cokurtosis"

TEMP_MATRIX(10, 1) = "t-value a1"
TEMP_MATRIX(11, 1) = "t-value a2"
TEMP_MATRIX(12, 1) = "t-value a3"
TEMP_MATRIX(13, 1) = "t-value a4"

TEMP_MATRIX(14, 1) = "Significance @ " & Format(1 - CI_FACTOR, "0.0%") & " a1"
TEMP_MATRIX(15, 1) = "Significance @ " & Format(1 - CI_FACTOR, "0.0%") & " a2"
TEMP_MATRIX(16, 1) = "Significance @ " & Format(1 - CI_FACTOR, "0.0%") & " a3"
TEMP_MATRIX(17, 1) = "Significance @ " & Format(1 - CI_FACTOR, "0.0%") & " a4"

TEMP_MATRIX(18, 1) = "a3 is significant & negative"
TEMP_MATRIX(19, 1) = "a4 is significant & negative"

TEMP_VAL = NORMSINV_FUNC(1 - CI_FACTOR / 2, 0, 1, 0)

For j = 1 To NCOLUMNS
    
    TEMP1_VECTOR = MATRIX_GET_COLUMN_FUNC(DATA_MATRIX, j, 1)
    TEMP2_VECTOR = DATA_BASIC_MOMENTS_FUNC(TEMP1_VECTOR, 0, 0, CI_FACTOR, 0)
    
    TEMP_MATRIX(1, j + 1) = TEMP2_VECTOR(1, 12)
    TEMP_MATRIX(2, j + 1) = TEMP2_VECTOR(1, 13)
    TEMP_MATRIX(3, j + 1) = JARQUE_BERA_HYPOTHESIS_MODEL_FUNC(TEMP2_VECTOR(1, 12), TEMP2_VECTOR(1, 13), TEMP2_VECTOR(1, 1), CI_FACTOR, 0)
    TEMP_MATRIX(4, j + 1) = JARQUE_BERA_HYPOTHESIS_MODEL_FUNC(TEMP2_VECTOR(1, 12), TEMP2_VECTOR(1, 13), TEMP2_VECTOR(1, 1), CI_FACTOR, 2)
    TEMP_MATRIX(5, j + 1) = REGRESSION_MOMENTS_INDEX_FUNC(TEMP1_VECTOR, XDATA_VECTOR, 0)
    TEMP_GROUP = REGRESSION_MOMENTS_INDEX_FUNC(TEMP1_VECTOR, XDATA_VECTOR, 9)
    
    TEMP_MATRIX(6, j + 1) = TEMP_GROUP(1)
    TEMP_MATRIX(7, j + 1) = TEMP_GROUP(2)
    TEMP_MATRIX(8, j + 1) = TEMP_GROUP(3)
    TEMP_MATRIX(9, j + 1) = TEMP_GROUP(4)
    TEMP_MATRIX(10, j + 1) = TEMP_GROUP(5)
    TEMP_MATRIX(11, j + 1) = TEMP_GROUP(6)
    TEMP_MATRIX(12, j + 1) = TEMP_GROUP(7)
    TEMP_MATRIX(13, j + 1) = TEMP_GROUP(8)
    
    TEMP_MATRIX(14, j + 1) = Abs(TEMP_MATRIX(10, j + 1)) >= TEMP_VAL
    TEMP_MATRIX(15, j + 1) = Abs(TEMP_MATRIX(11, j + 1)) >= TEMP_VAL
    TEMP_MATRIX(16, j + 1) = Abs(TEMP_MATRIX(12, j + 1)) >= TEMP_VAL
    TEMP_MATRIX(17, j + 1) = Abs(TEMP_MATRIX(13, j + 1)) >= TEMP_VAL

    TEMP_MATRIX(18, j + 1) = IIf((TEMP_MATRIX(12, j + 1) < 0) And (TEMP_MATRIX(16, j + 1) = True), True, False)
    TEMP_MATRIX(19, j + 1) = IIf(((TEMP_MATRIX(13, j + 1) < 0) And (TEMP_MATRIX(17, j + 1) = True)), True, False)
Next j

Select Case OUTPUT
Case 0
    FOUR_MOMENTS_INDEX_FUNC = MATRIX_TRANSPOSE_FUNC(TEMP_MATRIX)
Case Else
    FOUR_MOMENTS_INDEX_FUNC = TEMP_MATRIX
End Select

Exit Function
ERROR_LABEL:
FOUR_MOMENTS_INDEX_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : REGRESSION_MOMENTS_INDEX_FUNC
'DESCRIPTION   :
'LIBRARY       : STATISTICS
'GROUP         : MOMENTS
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 14/02/2008
'************************************************************************************
'************************************************************************************

Function REGRESSION_MOMENTS_INDEX_FUNC(ByRef YDATA_RNG As Variant, _
ByRef XDATA_RNG As Variant, _
Optional ByVal OUTPUT As Integer = 0)

'Skewness, kurtosis, tail measures and distribution tests are
'statistical approaches to describe the characteristics of a
'given return distribution.

'A more economic approach is the Four-Moment Index Model (also called
'the cubic model, four-moment CAPM, depending on the context), which
'extends the single-index model model by including squared and cubic
'unexpected index returns as additional factors.

'The cubic model is defined as follows...
'r(i,t) - r(f,t) = a1 + a2*{r(m,t)-r(f,t)} +
'a3*{r(m,t)-mean[r(m)]}^2 + a4*{r(m,t)-mean[r(m)]}^3 + e(t)
'with...
'r(i,t)... return of instrument i at point in time t
'r(f,t)... riskfree rate in t
'r(m,t)... benchmark return in t
'mean[r(m)]... expected indexc return calculated as the arithmetic mean return
'r(m,t)-mean[r(m)]... unexpected index return
'e(t).... white noise

'The most important goal of the cubic model is to test which of the parameters
'a1, a2, a3 are significantly different from zero (=t-tests). The parameters
'can be interpreted as sensitivities relative to the benchmark skewness and
'kurtosis (co-skewness and co-kurtosis).
 

'The risk premium of a non-normally distributed asset is equal to:
'- its market risk multiplied the market premium b1 plus
'- its systematic skewness risk multiplied with the systematic skewness
'  market premium b2 plus
'- its systematic kurtosis risk multiplied with the systematic kurtosis
'  market premium b3

'If we assume that it is possible to construct a portfolio with zero BETA, zero
'systematic kurtosis, and unitary systematic skewness, then the market premium of
'this portfolio will be b2. If we assume that it is possible to construct a
'portfolio with zero BETA, zero systematic skewness, and unitary systematic
'kurtosis, then the market premium of this portfolio will be b3.


'We see that an asset required rate of return is composed of the risk free rate
'plus three premiums:
'- the first premium is the reward of having in the portfolio an asset which
'is contributing positively to the world market BETA
'- the second premium is the reward of having in the portfolio an asset which
'is contributing negatively to the world market skewness
'- the fourth premium is the reward of having in the portfolio an asset which
'is contributing positively to the world market kurtosis.


Dim i As Long
Dim NROWS As Long

Dim MEAN_VAL As Double
Dim YDATA_VECTOR As Variant
Dim XDATA_VECTOR As Variant

Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

XDATA_VECTOR = XDATA_RNG
If UBound(XDATA_VECTOR, 1) = 1 Then
    XDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(XDATA_VECTOR)
End If

YDATA_VECTOR = YDATA_RNG
If UBound(YDATA_VECTOR, 1) = 1 Then
    YDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(YDATA_VECTOR)
End If
If UBound(XDATA_VECTOR, 1) <> UBound(YDATA_VECTOR, 1) Then: GoTo ERROR_LABEL

NROWS = UBound(XDATA_VECTOR, 1)
MEAN_VAL = 0
For i = 1 To NROWS
    MEAN_VAL = MEAN_VAL + XDATA_VECTOR(i, 1)
Next i
MEAN_VAL = MEAN_VAL / NROWS

'--------------------------------------------------------------------------------
If OUTPUT = 0 Then
'--------------------------------------------------------------------------------
    ReDim TEMP_MATRIX(1 To NROWS, 1 To 4) ' build factors
    For i = 1 To NROWS
        TEMP_MATRIX(i, 1) = 1 ' Constant
        TEMP_MATRIX(i, 2) = XDATA_VECTOR(i, 1) ' Index factor
        TEMP_MATRIX(i, 3) = (XDATA_VECTOR(i, 1) - MEAN_VAL) ^ 2 ' Coskewness factor
        TEMP_MATRIX(i, 4) = (XDATA_VECTOR(i, 1) - MEAN_VAL) ^ 3 ' Cokurtosis factor
    Next i
    REGRESSION_MOMENTS_INDEX_FUNC = REGRESSION_LS2_FUNC(TEMP_MATRIX, YDATA_VECTOR, False, 0, 0.95, 1)
'--------------------------------------------------------------------------------
Else
'--------------------------------------------------------------------------------
    ReDim TEMP_MATRIX(1 To NROWS, 1 To 3)
    For i = 1 To NROWS
        TEMP_MATRIX(i, 1) = XDATA_VECTOR(i, 1) ' Index factor
        TEMP_MATRIX(i, 2) = (XDATA_VECTOR(i, 1) - MEAN_VAL) ^ 2 ' Coskewness factor
        TEMP_MATRIX(i, 3) = (XDATA_VECTOR(i, 1) - MEAN_VAL) ^ 3 ' Cokurtosis factor
    Next i
    REGRESSION_MOMENTS_INDEX_FUNC = REGRESSION_LS2_FUNC(TEMP_MATRIX, YDATA_VECTOR, True, 0, 0.95, 1)
'--------------------------------------------------------------------------------
End If
'--------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
REGRESSION_MOMENTS_INDEX_FUNC = Err.number
End Function
