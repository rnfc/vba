Attribute VB_Name = "FINAN_PORT_SIMUL_SAMPLING_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
    

'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_RETURNS_SAMPLING_FUNC
'DESCRIPTION   : Random Sampling of Returns
'LIBRARY       : FINAN_PORT
'GROUP         : SIMUL_SAMPLING
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************

Function PORT_RETURNS_SAMPLING_FUNC(ByRef DATA_RNG As Variant, _
ByVal THRESHOLD As Double, _
ByVal PERIODS As Long, _
ByVal nLOOPS As Long, _
ByVal MIN_PRICE As Double, _
ByVal DELTA_PRICE As Double, _
Optional ByVal NBINS As Long = 20, _
Optional ByVal DATA_TYPE As Integer = 1, _
Optional ByVal LOG_SCALE As Integer = 0)

Dim i As Long
Dim j As Long
Dim k As Long

Dim TEMP_VAL As Double
Dim RND_VALUE As Double

Dim TEMP_VECTOR As Variant
Dim TEMP_MATRIX As Variant
Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then: _
DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
If DATA_TYPE <> 0 Then DATA_VECTOR = MATRIX_PERCENT_FUNC(DATA_VECTOR, LOG_SCALE)

ReDim TEMP_MATRIX(1 To NBINS + 1, 1 To 2)

TEMP_MATRIX(1, 1) = MIN_PRICE
For i = 2 To NBINS + 1
    TEMP_MATRIX(i, 1) = TEMP_MATRIX(1, 1) + DELTA_PRICE * (i - 1)
Next i

For j = 1 To nLOOPS
    
    TEMP_VAL = THRESHOLD
    
    TEMP_VECTOR = RANDOM_SAMPLING_CHOOSER_FUNC(DATA_VECTOR, PERIODS, 1, , True)
    For i = 1 To PERIODS
        RND_VALUE = TEMP_VECTOR(i, 1)
        If LOG_SCALE = 1 Then
            TEMP_VAL = TEMP_VAL * Exp(RND_VALUE)
        Else: TEMP_VAL = TEMP_VAL * (1 + RND_VALUE)
        End If
    Next i

    For k = 1 To NBINS
        If (TEMP_VAL >= TEMP_MATRIX(k, 1)) And _
        (TEMP_VAL < TEMP_MATRIX(k + 1, 1)) Then: _
        TEMP_MATRIX(k, 2) = TEMP_MATRIX(k, 2) + 1
    Next k

Next j

PORT_RETURNS_SAMPLING_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
PORT_RETURNS_SAMPLING_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_RETURNS_SAMPLING_FUNC
'DESCRIPTION   : Random Sampling of Port Return (Normal Distribution)
'LIBRARY       : FINAN_PORT
'GROUP         : SIMUL_SAMPLING
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************

Function PORT_PERIOD_RETURN_SAMPLING_FUNC(Optional ByVal EXPECTED_RETURN As Double = 15, _
Optional ByVal VOLATILITY As Double = 20, _
Optional ByVal RANDOM_VAL As Double = 0, _
Optional ByVal COUNT_BASIS As Double = 365)

On Error GoTo ERROR_LABEL

If RANDOM_VAL = 0 Then: RANDOM_VAL = Rnd
PORT_PERIOD_RETURN_SAMPLING_FUNC = EXPECTED_RETURN * (1 / COUNT_BASIS) + _
                               NORMSINV_FUNC(RANDOM_VAL, 0, 1, 0) * _
                               VOLATILITY * Sqr(1 / COUNT_BASIS)

Exit Function
ERROR_LABEL:
PORT_PERIOD_RETURN_SAMPLING_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_RETURNS_BOOTSTRAP1_FUNC
'DESCRIPTION   : Boostrapping port returns

'…if the resample period length equals the datafrequency, then the
'PORT_RETURNS_BOOTSTRAP1_FUNC function preserves skewness and kurtosis
'characteristics

'…the larger the difference between the data frequency and the resample
'period length (for example, a value of 1 versus 250), the
'stronger the tendency for the bootstrapped series to converge to the
'normal distribution (this is the central limit theorem at work)

'…additionally, this bootstrap function does not account for
'autocorrelation effects like for example volatility clustering

'LIBRARY       : FINAN_PORT
'GROUP         : SIMUL_SAMPLING
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************

Function PORT_RETURNS_BOOTSTRAP1_FUNC(ByRef DATA_RNG As Variant, _
Optional ByRef NSIZE As Long = 0, _
Optional ByVal COUNT_BASIS As Double = 30, _
Optional ByVal DATA_TYPE As Integer = 0, _
Optional ByVal LOG_SCALE As Integer = 0, _
Optional ByVal OUTPUT As Integer = 0)

'COUNT BASIS you would like to resample. Input data frequency is daily.
'So if you want to resample monthly returns, the perdiod length would be
'30. To resample annual returns, choose 250.

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
If DATA_TYPE <> 0 Then DATA_MATRIX = MATRIX_PERCENT_FUNC(DATA_MATRIX, LOG_SCALE)

NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)
If (NSIZE = 0) Or (NSIZE > NROWS) Then: NSIZE = NROWS

ReDim TEMP_MATRIX(1 To NSIZE, 1 To NCOLUMNS)

For i = 1 To NSIZE
    For k = 1 To COUNT_BASIS
        h = Int(Rnd() * NROWS) + 1
        For j = 1 To NCOLUMNS
            TEMP_MATRIX(i, j) = TEMP_MATRIX(i, j) + DATA_MATRIX(h, j)
        Next j
    Next k
Next i

'---------------------------------------------------------------------------------
Select Case OUTPUT
'---------------------------------------------------------------------------------
Case 0
'---------------------------------------------------------------------------------
    PORT_RETURNS_BOOTSTRAP1_FUNC = TEMP_MATRIX
'---------------------------------------------------------------------------------
Case 1
'---------------------------------------------------------------------------------
'The 3rd & 4th moment (SKEW & KURTOSIS) is only preserved when the resample
'period length equals 1 (i.e. is equal to the data frequency period)
'---------------------------------------------------------------------------------
    PORT_RETURNS_BOOTSTRAP1_FUNC = DATA_BASIC_MOMENTS_FUNC(DATA_MATRIX, 0, 0, 0.05, 1)
'---------------------------------------------------------------------------------
Case Else
'---------------------------------------------------------------------------------
    PORT_RETURNS_BOOTSTRAP1_FUNC = MATRIX_CORRELATION_RANK_FUNC(TEMP_MATRIX, 0, 0)
'---------------------------------------------------------------------------------
End Select
'---------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
PORT_RETURNS_BOOTSTRAP1_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_RETURNS_BOOTSTRAP2_FUNC

'DESCRIPTION   : After playing with the Central Limit Theorem I got to thinking that ...
'Once upon a time I was asked by John Bollinger about the relationship between the Standard
'Deviation of daily stock returns and the Standard Deviation of stock prices over the past n
'days. When one speaks of the Standard Deviation (as it concerns stocks), one (usually) is
'referring to the SD of returns, not prices. However, Bollinger Bands considers the SD of
'prices ... which ain't usual. Anyway, I tried to find a relationship and wrote a (lousy)
'tutorial on the subject. Now I figure I should look at the daily returns over some n-day
'time period and calculate the total return over those n days and forget any mathematical
'gesticulation.

'Suppose the daily returns for a month are r1, r2 ... r25 ... where we pretend that there
'are n = 25 market days in a month. We calculate the total n-day gain: (1+r1)(1+r2)...(1+r25).
'By gain" I mean that $1 will become $(1+r1)(1+r2)...(1+r25) after n = 25 days.
'Okay, so I do this (to see what the distribution of Total n-day Gains might look like):
'I look at the daily returns for GE stock over the past 10 years. That's (about) 2500 daily returns.
'I pick 25 successive returns at random and calculate the 25-day gain.
'I repeat this ritual a jillion times and plot the distribution of these monthly gains.

'Yes, but I'm not interested in the distribution. I just like to see that it looks like a "bell"
'... not necessarily a "Normal" bell. If it's bell-shaped, we'll feel warm all over.
'It shows the possibilities for a $1K portfolio after n-days.
'I tried 25-days then 50-days then 75-days. (That's like 1, 2 and 3 months worth of daily returns)
'The interesting thing (for me) is that I used all daily stock return info so that ...

'Look at a jillion daily historical returns and extract a Mean and Standard Deviation
'... and whatever other numbers seem tasty.
'Discard the jillion returns, retain the magic numbers and construct some distribution.
'Then use that distribution to predict future portfolios.

'If you use actual returns, then you don't have to assume some distribution, right?
'Right ... but there ain't no interesting math involved.

'I remember getting all excited about Ito Calculus and writing an umpteen part tutorial.
'It was really neat and, at the end, I could provide portfolio dsitributions T years into the
'future, like this of course, it 's a log-norml distribution defined by a Mean and Standard
'Deviation and ...

'And it's smooooth, not your jagged distribution. I like it!
'Well, predicting the future is a black art so you might as well use whatever makes you happy.
'Using historical returns is like saying: "I expect the future to be something like the past."

'The thing that I like about generating charts like this one is this: It's Exact!
'i 'm not claiming that it's some future distribution. I'm claiming that it's what actually
'happened in the past. Then you should use whatever makes you happy.
'Yes, of course. It's entertaining, fascinating, educational, time consuming ...

'LIBRARY       : FINAN_PORT
'GROUP         : SIMUL_SAMPLING
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'REFERENCE:
'http://www.gummy-stuff.org/stock-returns.htm
'http://www.gummy-stuff.org/central-limit-theorem.htm
'************************************************************************************
'************************************************************************************

Function PORT_RETURNS_BOOTSTRAP2_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal DATA_TYPE As Integer = 1, _
Optional ByVal LOG_SCALE As Integer = 0, _
Optional ByVal FACTOR_VAL As Double = 1000, _
Optional ByVal NO_PERIODS As Long = 25, _
Optional ByVal nLOOPS As Long = 10000, _
Optional ByVal OUTPUT As Integer = 0)

'FACTOR_VAL -> INITIAL PORTFOLIO
'NO_PERIODS -> Type in the number of periods in the time period.
'nLOOPS -> Type in the number of n-period gains you want to calculate.

Dim i As Long
Dim j As Long
Dim k As Long

Dim SROW As Long 'start row
Dim NROWS As Long 'end row
Dim NBINS As Long

Dim MIN_VAL As Double
Dim MAX_VAL As Double

Dim BIN_MIN As Double
Dim BIN_WIDTH As Double
Dim GAIN_VAL As Double

Dim DATA_VECTOR As Variant
Dim TEMP_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then
    DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
End If
If DATA_TYPE <> 0 Then: DATA_VECTOR = MATRIX_PERCENT_FUNC(DATA_VECTOR, LOG_SCALE)
k = UBound(DATA_VECTOR, 1) - NO_PERIODS
If k <= 0 Then: GoTo ERROR_LABEL

Randomize
ReDim TEMP_VECTOR(1 To nLOOPS, 1 To 1)

MIN_VAL = 2 ^ 52: MAX_VAL = -2 ^ 52
For i = 1 To nLOOPS 'No Periods (day) - gains
    SROW = Int(k * Rnd) + 1 'Int(0.99999999 * 2515) = 2514
    NROWS = SROW + NO_PERIODS
    GAIN_VAL = 1
    For j = SROW To NROWS: GAIN_VAL = GAIN_VAL * (DATA_VECTOR(j, 1) + 1): Next j
    TEMP_VECTOR(i, 1) = GAIN_VAL
    If GAIN_VAL < MIN_VAL Then: MIN_VAL = GAIN_VAL
    If GAIN_VAL > MAX_VAL Then: MAX_VAL = GAIN_VAL
Next i

If OUTPUT <> 0 Then
    PORT_RETURNS_BOOTSTRAP2_FUNC = TEMP_VECTOR
    Exit Function
End If

'DATA_VECTOR = DATA_BASIC_MOMENTS_FUNC(TEMP_VECTOR, 0, 0, 0.05, 0)
'DATA_VECTOR = HISTOGRAM_BIN_LIMITS_FUNC(DATA_VECTOR(1, 2), DATA_VECTOR(1, 3), nLOOPS, 3)
DATA_VECTOR = HISTOGRAM_BIN_LIMITS_FUNC(MIN_VAL, MAX_VAL, nLOOPS, 3)
BIN_WIDTH = DATA_VECTOR(LBound(DATA_VECTOR))
BIN_MIN = DATA_VECTOR(LBound(DATA_VECTOR) + 1)
NBINS = DATA_VECTOR(LBound(DATA_VECTOR) + 2)
DATA_VECTOR = HISTOGRAM_FREQUENCY_FUNC(TEMP_VECTOR, NBINS, BIN_MIN, BIN_WIDTH, 1)
NBINS = UBound(DATA_VECTOR, 1)

ReDim TEMP_VECTOR(0 To NBINS, 1 To 5)
TEMP_VECTOR(0, 1) = "PORTFOLIO"
TEMP_VECTOR(0, 2) = "BINS"
TEMP_VECTOR(0, 3) = "FREQ"
TEMP_VECTOR(0, 4) = "PDF: " & NO_PERIODS & " GAINS"
TEMP_VECTOR(0, 5) = "CDF: " & NO_PERIODS & " GAINS"

For i = 1 To NBINS
    TEMP_VECTOR(i, 1) = DATA_VECTOR(i, 1) * FACTOR_VAL
    TEMP_VECTOR(i, 2) = DATA_VECTOR(i, 1)
    TEMP_VECTOR(i, 3) = DATA_VECTOR(i, 2)
    TEMP_VECTOR(i, 4) = TEMP_VECTOR(i, 3) / nLOOPS
    If i > 1 Then
        TEMP_VECTOR(i, 5) = TEMP_VECTOR(i - 1, 5) + TEMP_VECTOR(i, 4)
    Else
        TEMP_VECTOR(i, 5) = TEMP_VECTOR(i, 4)
    End If
Next i

PORT_RETURNS_BOOTSTRAP2_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
PORT_RETURNS_BOOTSTRAP2_FUNC = Err.number
End Function
