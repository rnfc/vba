Attribute VB_Name = "FINAN_ASSET_SYSTEM_KELLY_LIBR"

'--------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------
Option Explicit
Option Base 1
'--------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------
'the Kelly Ratio
'---------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------
'In 1956, John Kelly* at AT&T's Bell Labs did research on telephone transmission in the presence
'of noise and ... We 'll just talk about the results of his analysis as it has come to be applied
'to the stock market (called, among other things, Kelly Ratio/Value/Criterion), in particular,
'what it says about how much money to put into a single trade, given the historical evolution of
'the stock: the percentage of times that you'd win, the average winnings per trade compared to
'the average loss per trade and ...

'Kelly% = the percentage of your capital to be put into a single trade
'Note: For a p = 50% probability of winning (like tossing a coin),
'and equal winnings as losses (so W = L), Kelly says
'.. but if you expect to win p = 60% of the time, then Kelly says
'Put 20% of your capital into your next trade

'Well , Kelly 's original paper is quite mathematical ...
'... but we can explain the formula like so:
'* When you make a winning trade, you average $500.
'* Your losses, per trade, average $350.
'* The probability of winning (and making, on average, $500) is 0.60 ... the Winning Probability.
'* Out of 1000 trades, you'd expect to win 0.60 *1000 = 600 times and lose (1-0.60) * 1000 = 400 times.
'* The wins provide (600)*500 = $300,000 and the total expected losses are (400)*350 = $140,000
'... for 1000 trades.
'* Hence, the expected gain is $300,000 - $140,000 = $160,000 ... for 1000 trades.
'* The expected gain per trade is $160 ... dividing by 1000.
'* Then you expect to make this amount per trade ... on average.
'* Then these expected winnings of $160 is just 160/500 = 0.32 or 32% of your winning trades.
'* That's the Kelly Ratio!

'In general:
'* When you make a winning trade, you average $W.
'* Your losses, per trade, average $L.
'* The probability of winning (and making, on average, $W) is p ... the Winning Probability.
'* Out of N trades, you'd expect to win p *N times and lose (1-p) * N times.
'* The wins provide $(p*N)*W and the total expected losses are $[(1-p)*N]*L ... for N trades.
'* Hence, the expected gain is (p*N)*W - [(1-p)*N]*L ... for N trades.
'* The expected gain per trade is (p)*W - (1-p)*L ... dividing by N.
'* As a fraction of W, we get:     Kelly Ratio = { p*W - (1-p)*L } / W

'But it's interpreted as the percentage of your capital to invest in each trade. Why?
'If the expected Gain per trade is p*W - (1-p)*L and you'd like to make your winning gain,
'namely $W, then you'd have to make W/{ p*W - (1-p)*L } trades so you have to have enough money
'to make these trades so if this number was 4 then you'd invest just 1/4 or 25% of your capital
'on each trade so you'd have enough money to make four trades so you should only invest a fraction
'Kelly = { p*W - (1-p)*L }/W on each trade or you may just want to average $W over the long haul or maybe ...

'My sentiments exactly ... but not everybuddy uses the same formula for their Kelly Criterion, although
'the expression p*W - (1-p)*L seems to be in everybuddy's Kelly Criterion. That's your expected winnings
'per bet. For application to the stock market, and the use of the above formula, see
'http://www.hquotes.com/kelly.html


'In 1961, Kelly was involved in making a computer sing "A Bicycle Built for Two". Arthur C. Clark heard
'the computer-synthesized song when he visited the labs and had Hal the computer sing it in "2001: A
'Space Odyssey" ... when Hal was being disconnected.


'REFERENCES:
'http://www.gummy-stuff.org/kelly-ratio.htm
'http://www.elitetrader.com/vb/showthread.php?threadid=170543

'http://www.hquotes.com/kelly.html
'http://www.hquotes.com/charts/zerosum.pdf
'http://www.hquotes.com/tradehard/simulator.html

'http://www.bjmath.com/bjmath/thorp/tog.htm
'http://www.bjmath.com/bjmath/thorp/paper.htm

'http://www.racing.saratoga.ny.us/kelly.pdf
Function ASSET_KELLY_RATIO_FUNC(ByRef TICKER_STR As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal END_DATE As Date, _
Optional ByVal INITIAL_WEALTH As Double = 1000, _
Optional ByVal OUTPUT As Integer = 1)

Dim i As Long
Dim j As Long

Dim ii As Long 'Total Wins
Dim jj As Long 'Total Losses

Dim NROWS As Long

Dim PP_VAL As Double 'percentage of wins
Dim QQ_VAL As Double 'percentage of losses

Dim WIN_VAL As Double 'average of winning returns
Dim LOSS_VAL As Double 'average of losing returns

Dim SD_WIN_VAL As Double 'Standard Deviation of winning returns
Dim SD_LOSS_VAL As Double 'Standard Deviation of losing returns

Dim K1_VAL As Double 'Kelly #1: p-q/(W/L)
Dim K2_VAL As Double 'Kelly #2:(p*W-q*L)/(W*L)
Dim K3_VAL As Double 'Kelly #3:(p*W-q*L)/(p*(W^2+SD_win^2)+q*(L^2+SD_loss^2))
Dim BUY_HOLD_VAL As Double

Dim K1_MULT_VAL As Double
Dim K2_MULT_VAL As Double
Dim K3_MULT_VAL As Double
Dim BUY_HOLD_MULT_VAL As Double

Dim DATA_MATRIX As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

If IsArray(TICKER_STR) = True Then
    DATA_MATRIX = TICKER_STR
Else
    DATA_MATRIX = _
    YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, START_DATE, END_DATE, _
    "d", "DOHLCVA", False, True, True)
End If
NROWS = UBound(DATA_MATRIX, 1)

'--------------------------------------------------------------------------------
ReDim TEMP_MATRIX(0 To NROWS, 1 To 20)
'--------------------------------------------------------------------------------
TEMP_MATRIX(0, 1) = "DATE"
TEMP_MATRIX(0, 2) = "OPEN"
TEMP_MATRIX(0, 3) = "HIGH"
TEMP_MATRIX(0, 4) = "LOW"
TEMP_MATRIX(0, 5) = "CLOSE"
TEMP_MATRIX(0, 6) = "VOLUME"
TEMP_MATRIX(0, 7) = "ADJ CLOSE"
'--------------------------------------------------------------------------------
TEMP_MATRIX(0, 8) = "RETURNS"
TEMP_MATRIX(0, 9) = "WINS"
TEMP_MATRIX(0, 10) = "LOSSES"
'--------------------------------------------------------------------------------
TEMP_MATRIX(0, 11) = "CONSTANT WINS"
TEMP_MATRIX(0, 12) = "CONSTANT LOSSES"
'--------------------------------------------------------------------------------
TEMP_MATRIX(0, 13) = "GROWTH K1"
TEMP_MATRIX(0, 14) = "GROWTH K2"
TEMP_MATRIX(0, 15) = "GROWTH K3"
TEMP_MATRIX(0, 16) = "GROWTH B&H"
'--------------------------------------------------------------------------------
TEMP_MATRIX(0, 17) = "RETURNS K1"
TEMP_MATRIX(0, 18) = "RETURNS K2"
TEMP_MATRIX(0, 19) = "RETURNS K3"
TEMP_MATRIX(0, 20) = "RETURNS B&H"
'--------------------------------------------------------------------------------

ii = 0: jj = 0
WIN_VAL = 0: LOSS_VAL = 0

For i = 1 To NROWS
    For j = 1 To 7: TEMP_MATRIX(i, j) = DATA_MATRIX(i, j): Next j
    TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 6) / 1000
    
    TEMP_MATRIX(i, 8) = TEMP_MATRIX(i, 5) / TEMP_MATRIX(i, 2) - 1
    
    If TEMP_MATRIX(i, 8) > 0 Then
        TEMP_MATRIX(i, 9) = TEMP_MATRIX(i, 8)
        WIN_VAL = WIN_VAL + TEMP_MATRIX(i, 9)
        ii = ii + 1
        TEMP_MATRIX(i, 10) = ""
    ElseIf TEMP_MATRIX(i, 8) < 0 Then
        TEMP_MATRIX(i, 9) = ""
        TEMP_MATRIX(i, 10) = TEMP_MATRIX(i, 8)
        LOSS_VAL = LOSS_VAL + TEMP_MATRIX(i, 10)
        jj = jj + 1
    End If
Next i
WIN_VAL = WIN_VAL / ii
LOSS_VAL = LOSS_VAL / jj

PP_VAL = ii / (ii + jj)
QQ_VAL = 1 - ii / (ii + jj)

SD_WIN_VAL = 0
SD_LOSS_VAL = 0
For i = 1 To NROWS
    If TEMP_MATRIX(i, 8) > 0 Then
        SD_WIN_VAL = SD_WIN_VAL + (TEMP_MATRIX(i, 9) - WIN_VAL) ^ 2
    ElseIf TEMP_MATRIX(i, 8) < 0 Then
        SD_LOSS_VAL = SD_LOSS_VAL + (TEMP_MATRIX(i, 10) - LOSS_VAL) ^ 2
    End If
Next i
SD_WIN_VAL = (SD_WIN_VAL / ii) ^ 0.5
SD_LOSS_VAL = (SD_LOSS_VAL / jj) ^ 0.5

LOSS_VAL = Abs(LOSS_VAL)

K1_VAL = PP_VAL - QQ_VAL / (WIN_VAL / LOSS_VAL)
K2_VAL = (PP_VAL * WIN_VAL - QQ_VAL * LOSS_VAL) / (WIN_VAL * LOSS_VAL)
K3_VAL = (PP_VAL * WIN_VAL - QQ_VAL * LOSS_VAL) / (PP_VAL * (WIN_VAL ^ 2 + _
         SD_WIN_VAL ^ 2) + QQ_VAL * (LOSS_VAL ^ 2 + SD_LOSS_VAL ^ 2))
BUY_HOLD_VAL = 1 '100%

K1_MULT_VAL = INITIAL_WEALTH
K2_MULT_VAL = INITIAL_WEALTH
K3_MULT_VAL = INITIAL_WEALTH
BUY_HOLD_MULT_VAL = INITIAL_WEALTH
For i = 1 To NROWS
    TEMP_MATRIX(i, 11) = IIf(TEMP_MATRIX(i, 9) <> "", WIN_VAL, "")
    TEMP_MATRIX(i, 12) = IIf(TEMP_MATRIX(i, 10) <> "", LOSS_VAL, "")
    
    TEMP_MATRIX(i, 13) = K1_MULT_VAL * (1 + K1_VAL * TEMP_MATRIX(i, 8))
    TEMP_MATRIX(i, 14) = K2_MULT_VAL * (1 + K2_VAL * TEMP_MATRIX(i, 8))
    TEMP_MATRIX(i, 15) = K3_MULT_VAL * (1 + K3_VAL * TEMP_MATRIX(i, 8))
    TEMP_MATRIX(i, 16) = BUY_HOLD_MULT_VAL * (1 + BUY_HOLD_VAL * TEMP_MATRIX(i, 8))

    TEMP_MATRIX(i, 17) = TEMP_MATRIX(i, 13) / K1_MULT_VAL - 1
    TEMP_MATRIX(i, 18) = TEMP_MATRIX(i, 14) / K2_MULT_VAL - 1
    TEMP_MATRIX(i, 19) = TEMP_MATRIX(i, 15) / K3_MULT_VAL - 1
    TEMP_MATRIX(i, 20) = TEMP_MATRIX(i, 16) / BUY_HOLD_MULT_VAL - 1
    
    K1_MULT_VAL = TEMP_MATRIX(i, 13)
    K2_MULT_VAL = TEMP_MATRIX(i, 14)
    K3_MULT_VAL = TEMP_MATRIX(i, 15)
    BUY_HOLD_MULT_VAL = TEMP_MATRIX(i, 16)
Next i

Select Case OUTPUT
Case 0
    ASSET_KELLY_RATIO_FUNC = TEMP_MATRIX
Case Else
    ASSET_KELLY_RATIO_FUNC = Array(WIN_VAL, LOSS_VAL, ii, jj, PP_VAL, QQ_VAL, _
    SD_WIN_VAL, SD_LOSS_VAL, K1_VAL, K2_VAL, K3_VAL, BUY_HOLD_VAL, _
    K1_MULT_VAL, K2_MULT_VAL, K3_MULT_VAL, BUY_HOLD_MULT_VAL)
End Select

Exit Function
ERROR_LABEL:
ASSET_KELLY_RATIO_FUNC = Err.number
End Function

Function ASSETS_KELLY_RATIO_FUNC(ByRef TICKERS_RNG As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal END_DATE As Date, _
Optional ByVal INITIAL_WEALTH As Double = 1000)

Dim i As Long
Dim j As Long
Dim NROWS As Long
Dim TEMP_ARR As Variant
Dim TEMP_MATRIX As Variant
Dim TICKERS_VECTOR As Variant

On Error GoTo ERROR_LABEL

If IsArray(TICKERS_RNG) = True Then
    TICKERS_VECTOR = TICKERS_RNG
    If UBound(TICKERS_VECTOR, 1) = 1 Then
        TICKERS_VECTOR = MATRIX_TRANSPOSE_FUNC(TICKERS_VECTOR)
    End If
Else
    ReDim TICKERS_VECTOR(1 To 1, 1 To 1)
    TICKERS_VECTOR(1, 1) = TICKERS_RNG
End If
NROWS = UBound(TICKERS_VECTOR, 1)


ReDim TEMP_MATRIX(0 To NROWS, 1 To 19)
TEMP_MATRIX(0, 1) = "SYMBOL"
TEMP_MATRIX(0, 2) = "AVERAGE OF WINNING RETURNS (WIN)"
TEMP_MATRIX(0, 3) = "AVERAGE OF LOSING RETURNS (LOSS)"
TEMP_MATRIX(0, 4) = "WIN/LOSS"
TEMP_MATRIX(0, 5) = "TOTAL"
TEMP_MATRIX(0, 6) = "WINS"
TEMP_MATRIX(0, 7) = "LOSSES"
TEMP_MATRIX(0, 8) = "PERCENTAGE OF WINS (PP)"
TEMP_MATRIX(0, 9) = "PERCENTAGE OF LOSSES (QQ)"
TEMP_MATRIX(0, 10) = "STANDARD DEVIATION OF WINNING RETURNS (SD_WIN)"
TEMP_MATRIX(0, 11) = "STANDARD DEVIATION OF LOSING RETURNS (SD_LOSS)"
TEMP_MATRIX(0, 12) = "KELLY #1 (%) = P-Q/(W/L) "
TEMP_MATRIX(0, 13) = "KELLY #2 (%) = (P*W-Q*L)/(W*L)"
TEMP_MATRIX(0, 14) = "KELLY #3 (%) = (P*W-Q*L)/(P*(W^2+SD_WIN^2)+Q*(L^2+SD_LOSS^2))"
TEMP_MATRIX(0, 15) = "BUY & HOLD (%)"
TEMP_MATRIX(0, 16) = "KELLY #1 ($)"
TEMP_MATRIX(0, 17) = "KELLY #2 ($)"
TEMP_MATRIX(0, 18) = "KELLY #3 ($)"
TEMP_MATRIX(0, 19) = "BUY & HOLD ($)"

For i = 1 To NROWS
    TEMP_MATRIX(i, 1) = TICKERS_VECTOR(i, 1)
    TEMP_ARR = ASSET_KELLY_RATIO_FUNC(TICKERS_VECTOR(i, 1), START_DATE, END_DATE, INITIAL_WEALTH, 1)
    If IsArray(TEMP_ARR) = False Then: GoTo 1983
    j = LBound(TEMP_ARR)
    
    TEMP_MATRIX(i, 2) = TEMP_ARR(j)
    TEMP_MATRIX(i, 3) = TEMP_ARR(j + 1)
    If TEMP_MATRIX(i, 3) <> 0 Then
        TEMP_MATRIX(i, 4) = TEMP_MATRIX(i, 2) / TEMP_MATRIX(i, 3)
    Else
        TEMP_MATRIX(i, 4) = ""
    End If
    TEMP_MATRIX(i, 6) = TEMP_ARR(j + 2)
    TEMP_MATRIX(i, 7) = TEMP_ARR(j + 3)
    
    TEMP_MATRIX(i, 5) = TEMP_MATRIX(i, 6) + TEMP_MATRIX(i, 7)
    TEMP_MATRIX(i, 8) = TEMP_ARR(j + 4)
    TEMP_MATRIX(i, 9) = TEMP_ARR(j + 5)
    TEMP_MATRIX(i, 10) = TEMP_ARR(j + 6)
    TEMP_MATRIX(i, 11) = TEMP_ARR(j + 7)
    
    TEMP_MATRIX(i, 12) = TEMP_ARR(j + 8)
    TEMP_MATRIX(i, 13) = TEMP_ARR(j + 9)
    TEMP_MATRIX(i, 14) = TEMP_ARR(j + 10)
    TEMP_MATRIX(i, 15) = TEMP_ARR(j + 11)
    
    TEMP_MATRIX(i, 16) = TEMP_ARR(j + 12)
    TEMP_MATRIX(i, 17) = TEMP_ARR(j + 13)
    TEMP_MATRIX(i, 18) = TEMP_ARR(j + 14)
    TEMP_MATRIX(i, 19) = TEMP_ARR(j + 15)
1983:
Next i

ASSETS_KELLY_RATIO_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
ASSETS_KELLY_RATIO_FUNC = Err.number
End Function


Function KELLY_RATIO_CONSTANT_PARAMETERS_FUNC(ByVal INITIAL_WEALTH As Double, _
ByVal WIN_RETURN As Double, _
ByVal LOSS_RETURN As Double, _
ByVal SIGMA_WIN_RETURN As Double, _
ByVal SIGMA_LOSS_RETURN As Double, _
ByVal PROBABILITY_WIN As Double, _
Optional ByVal MIN_BIN As Double = 0, _
Optional ByVal MAX_BIN As Double = 1, _
Optional ByVal DELTA_BIN As Double = 0.1, _
Optional ByVal OUTPUT As Integer = 1)

'INITIAL_WEALTH: I = 1000 = Initial Wealth
'WIN_RETURN: W = 15.0%    = return when you WIN
'LOSS_RETURN: L = 10.0%    = return when you LOSE
'SIGMA_WIN_RETURN: S(W) =  20.0%    = Standard Deviation of WIN returns
'SIGMA_LOSS_RETURN: S(L) =  25.0%    = Standard Deviation of LOSS returns
'PROBABILITY_WIN: p = 45.0%    = probability of a WIN

Dim j As Long
Dim NCOLUMNS As Long

Dim X_VAL As Double

Dim K1_VAL As Double 'p-q/(W/L)
Dim K2_VAL As Double '(p*W-q*L)/(W*L)
Dim K3_VAL As Double '(p*W-q*L)/( p*(W^2+SD_win^2)+q*(L^2+SD_loss^2))

Dim N_VAL As Double 'wins out of 100 (follows from user-defined p, above)
Dim M_VAL As Double 'losses out of 100 (follows from user-defined p, above)
Dim Q_VAL As Double 'percentage of losing returns as per user value of p, above (probability of a LOSS)

Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

N_VAL = 100 * PROBABILITY_WIN
M_VAL = 100 - N_VAL

'-----------------------------------------------------------------------------
Select Case OUTPUT
'-----------------------------------------------------------------------------
Case 0
'-----------------------------------------------------------------------------
    
    Q_VAL = 1 - PROBABILITY_WIN
    
    K1_VAL = PROBABILITY_WIN - Q_VAL / (WIN_RETURN / LOSS_RETURN)
    K2_VAL = (PROBABILITY_WIN * WIN_RETURN - Q_VAL * LOSS_RETURN) / (WIN_RETURN * LOSS_RETURN)
    K3_VAL = (PROBABILITY_WIN * WIN_RETURN - Q_VAL * LOSS_RETURN) / (PROBABILITY_WIN * _
             (WIN_RETURN ^ 2 + SIGMA_WIN_RETURN ^ 2) + Q_VAL * (LOSS_RETURN ^ 2 + SIGMA_LOSS_RETURN ^ 2))
    
    ReDim TEMP_MATRIX(1 To 4, 1 To 3)

    TEMP_MATRIX(1, 1) = "Kelly Constant Win/Loss Returns & Volatility"
    TEMP_MATRIX(2, 1) = "Kelly #1 Portfolio"
    TEMP_MATRIX(3, 1) = "Kelly #2 Portfolio"
    TEMP_MATRIX(4, 1) = "Kelly #3 Portfolio"
    
    TEMP_MATRIX(1, 2) = "Probability"
    TEMP_MATRIX(2, 2) = K1_VAL
    TEMP_MATRIX(3, 2) = K2_VAL
    TEMP_MATRIX(4, 2) = K3_VAL
    
    TEMP_MATRIX(1, 3) = "Portfolio"
    
    TEMP_MATRIX(2, 3) = INITIAL_WEALTH * (1 + K1_VAL * WIN_RETURN) ^ N_VAL * _
                        (1 - K1_VAL * LOSS_RETURN) ^ M_VAL
    
    TEMP_MATRIX(3, 3) = INITIAL_WEALTH * (1 + K2_VAL * WIN_RETURN) ^ N_VAL * _
                        (1 - K2_VAL * LOSS_RETURN) ^ M_VAL
    
    TEMP_MATRIX(4, 3) = INITIAL_WEALTH * (1 + K3_VAL * WIN_RETURN) ^ N_VAL * _
                        (1 - K3_VAL * LOSS_RETURN) ^ M_VAL

'-----------------------------------------------------------------------------
Case Else
'-----------------------------------------------------------------------------
    NCOLUMNS = Int((MAX_BIN - MIN_BIN) / DELTA_BIN) + 1
    ReDim TEMP_MATRIX(1 To 4, 1 To NCOLUMNS + 1)
    TEMP_MATRIX(1, 1) = "POINT"
    TEMP_MATRIX(2, 1) = "WINNING"
    TEMP_MATRIX(3, 1) = "LOSSING"
    TEMP_MATRIX(4, 1) = "WEALTH"
    
    X_VAL = MIN_BIN
    For j = 1 To NCOLUMNS
        TEMP_MATRIX(1, j + 1) = X_VAL
        TEMP_MATRIX(2, j + 1) = (1 + X_VAL * WIN_RETURN) ^ N_VAL
        TEMP_MATRIX(3, j + 1) = (1 - X_VAL * LOSS_RETURN) ^ M_VAL
        TEMP_MATRIX(4, j + 1) = INITIAL_WEALTH * TEMP_MATRIX(2, j + 1) * TEMP_MATRIX(3, j + 1)
        
        X_VAL = X_VAL + DELTA_BIN
    Next j
'-----------------------------------------------------------------------------
End Select
'-----------------------------------------------------------------------------

KELLY_RATIO_CONSTANT_PARAMETERS_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
KELLY_RATIO_CONSTANT_PARAMETERS_FUNC = Err.number
End Function


'Theoretical Values for a Normal Distribution

Function KELLY_RATIO_NORMAL_TABLE_FUNC(ByVal MEAN_PER_PERIOD As Double, _
ByVal SIGMA_PER_PERIOD As Double, _
Optional ByVal NBINS As Long = 399, _
Optional ByVal OUTPUT As Integer = 1)

Dim i As Long
Dim j As Long

Dim Z_VAL As Double

Dim AVG_WIN_VAL As Double
Dim AVG_LOSS_VAL As Double

Dim SD_WIN_VAL As Double
Dim SD_LOSS_VAL As Double

Dim TEMP_SUM As Double
Dim DELTA_BIN As Double
Dim TEMP_MATRIX As Variant
Dim TEMP_VECTOR As Variant

On Error GoTo ERROR_LABEL

Z_VAL = 4
DELTA_BIN = 1 / (NBINS + 1)

ReDim TEMP_MATRIX(0 To NBINS, 1 To 4)
TEMP_MATRIX(0, 1) = "x"
TEMP_MATRIX(0, 2) = "x^2"
TEMP_MATRIX(0, 3) = "f(x)"
TEMP_MATRIX(0, 4) = "dummy"

j = 0
TEMP_SUM = 0
For i = 1 To NBINS
    If i <> 1 Then
        TEMP_MATRIX(i, 1) = TEMP_MATRIX(i - 1, 1) + Z_VAL * 2 * SIGMA_PER_PERIOD / (NBINS + 1)
    Else
        TEMP_MATRIX(i, 1) = MEAN_PER_PERIOD - Z_VAL * SIGMA_PER_PERIOD
    End If
    TEMP_MATRIX(i, 2) = TEMP_MATRIX(i, 1) ^ 2
    TEMP_MATRIX(i, 3) = NORMDIST_FUNC(TEMP_MATRIX(i, 1), MEAN_PER_PERIOD, SIGMA_PER_PERIOD, 0)
    TEMP_SUM = TEMP_SUM + TEMP_MATRIX(i, 3)
    
    TEMP_MATRIX(i, 4) = IIf(TEMP_MATRIX(i, 1) < 0, 1, 0)
    j = j + TEMP_MATRIX(i, 4)
Next i

If OUTPUT = 0 Then
    KELLY_RATIO_NORMAL_TABLE_FUNC = TEMP_MATRIX
    Exit Function
End If

AVG_WIN_VAL = 0: AVG_LOSS_VAL = 0
SD_WIN_VAL = 0: SD_LOSS_VAL = 0
For i = 1 To NBINS
    If i > j + 1 Then
        AVG_WIN_VAL = AVG_WIN_VAL + TEMP_MATRIX(i, 3) * TEMP_MATRIX(i, 1)
        SD_WIN_VAL = SD_WIN_VAL + TEMP_MATRIX(i, 3) * TEMP_MATRIX(i, 2)
    Else 'If i <= j Then:
        AVG_LOSS_VAL = AVG_LOSS_VAL + TEMP_MATRIX(i, 3) * TEMP_MATRIX(i, 1)
        SD_LOSS_VAL = SD_LOSS_VAL + TEMP_MATRIX(i, 3) * TEMP_MATRIX(i, 2)
    End If
Next i

AVG_WIN_VAL = AVG_WIN_VAL / TEMP_SUM
AVG_LOSS_VAL = AVG_LOSS_VAL / TEMP_SUM
SD_WIN_VAL = Sqr(SD_WIN_VAL / TEMP_SUM)
SD_LOSS_VAL = Sqr(SD_LOSS_VAL / TEMP_SUM)

ReDim TEMP_VECTOR(1 To 7, 1 To 2)
TEMP_VECTOR(1, 1) = "SUM(f(x))"
TEMP_VECTOR(2, 1) = "avg win"
TEMP_VECTOR(3, 1) = "f(x) win"
TEMP_VECTOR(4, 1) = "SD(win)"
TEMP_VECTOR(5, 1) = "avg loss"
TEMP_VECTOR(6, 1) = "f(x) loss"
TEMP_VECTOR(7, 1) = "SD(loss)"

TEMP_VECTOR(1, 2) = TEMP_SUM
TEMP_VECTOR(2, 2) = AVG_WIN_VAL
TEMP_VECTOR(3, 2) = NORMDIST_FUNC(AVG_WIN_VAL, MEAN_PER_PERIOD, SIGMA_PER_PERIOD, 0)
TEMP_VECTOR(4, 2) = SD_WIN_VAL
TEMP_VECTOR(5, 2) = AVG_LOSS_VAL
TEMP_VECTOR(6, 2) = NORMDIST_FUNC(AVG_LOSS_VAL, MEAN_PER_PERIOD, SIGMA_PER_PERIOD, 0)
TEMP_VECTOR(7, 2) = SD_LOSS_VAL

If OUTPUT = 1 Then
    KELLY_RATIO_NORMAL_TABLE_FUNC = TEMP_VECTOR
    Exit Function
End If

KELLY_RATIO_NORMAL_TABLE_FUNC = Array(TEMP_MATRIX, TEMP_VECTOR)

Exit Function
ERROR_LABEL:
KELLY_RATIO_NORMAL_TABLE_FUNC = Err.number
End Function

'Returns the avg win and avg loss returns & volatilities and a sample portfolio
'for each % of bankroll --> 25%, 50%, 75% and 100%

Function KELLY_RATIO_MC_FUNC(ByVal MEAN_PER_PERIOD As Double, _
ByVal SIGMA_PER_PERIOD As Double, _
ByVal K1_VAL As Double, _
ByVal K2_VAL As Double, _
ByVal K3_VAL As Double, _
ByVal K4_VAL As Double, _
Optional ByVal NBINS As Long = 399, _
Optional ByVal PERIODS As Long = 100, _
Optional ByVal nLOOPS As Long = 10000)

'Average Return (per period) =   4.0%
'Standard Deviation (per period) =   30.0%
'Number of periods = 100
'Number of Monte Carlo simulations = 500
'get number of periods
'get number of Monte Carlo iterations
'get four investment percentages (K1, K2, K3, K4) --> % of bankroll invested per trade
'25% 50% 75% 100%


Dim i As Long
Dim j As Long

Dim E_VAL As Double

Dim P1_VAL As Double
Dim P2_VAL As Double
Dim P3_VAL As Double
Dim P4_VAL As Double

Dim MEAN1_VAL As Double
Dim MEAN2_VAL As Double
Dim MEAN3_VAL As Double
Dim MEAN4_VAL As Double

Dim P_VAL As Double
Dim WIN_VAL As Double
Dim LOSS_VAL As Double

Dim GAIN_VAL As Double
Dim SD_WIN_VAL As Double
Dim SD_LOSS_VAL As Double

Dim TEMP_BIN As Double
Dim DELTA_BIN As Double
Dim NORMAL_ARR() As Double

On Error GoTo ERROR_LABEL

DELTA_BIN = 1 / (NBINS + 1)
ReDim NORMAL_ARR(1 To NBINS)
TEMP_BIN = DELTA_BIN
For i = 1 To NBINS
    NORMAL_ARR(i) = NORMSINV_FUNC(TEMP_BIN, 0, 1, 0)
    TEMP_BIN = TEMP_BIN + DELTA_BIN
Next i

' initialize
MEAN1_VAL = 0
MEAN2_VAL = 0
MEAN3_VAL = 0
MEAN4_VAL = 0
P_VAL = 0
WIN_VAL = 0
LOSS_VAL = 0
SD_WIN_VAL = 0
SD_LOSS_VAL = 0

'   change to lognormal parameters
'SIGMA_PER_PERIOD = Sqr(1 + (SIGMA_PER_PERIOD / (1 + MEAN_PER_PERIOD)) ^ 2)
'MEAN_PER_PERIOD = Log(1 + MEAN_PER_PERIOD) - SIGMA_PER_PERIOD ^ 2 / 2 - 1
Randomize

For i = 1 To nLOOPS
    P1_VAL = 1  ' start sample portfolios at $1.00
    P2_VAL = 1
    P3_VAL = 1
    P4_VAL = 1
    
    For j = 1 To PERIODS
        E_VAL = 1 + (NBINS - 1) * Rnd
        E_VAL = NORMAL_ARR(E_VAL)
        GAIN_VAL = MEAN_PER_PERIOD + E_VAL * SIGMA_PER_PERIOD       ' generate a random return
'        GAIN_VAL = Exp(GAIN_VAL) - 1
        P1_VAL = P1_VAL * (1 + K1_VAL * GAIN_VAL)          ' change sample portfolios
        P2_VAL = P2_VAL * (1 + K2_VAL * GAIN_VAL)
        P3_VAL = P3_VAL * (1 + K3_VAL * GAIN_VAL)
        P4_VAL = P4_VAL * (1 + K4_VAL * GAIN_VAL)
        If GAIN_VAL >= 0 Then       ' check if it's a WIN_VAL
            P_VAL = P_VAL + 1
            WIN_VAL = WIN_VAL + GAIN_VAL                  ' calculate average WIN_VAL return
            SD_WIN_VAL = SD_WIN_VAL + GAIN_VAL * GAIN_VAL       ' calculate standard deviation
        Else                    ' check if it's a LOSS_VAL
            LOSS_VAL = LOSS_VAL + GAIN_VAL                ' calculate average LOSS_VAL return
            SD_LOSS_VAL = SD_LOSS_VAL + GAIN_VAL * GAIN_VAL     ' calculate standard deviation
        End If
        'If i = 1 Then              ' display (sample) FINAL portfolios
        '    Cells(j + 1, 7) = P1_VAL
        '    Cells(j + 1, 8) = P2_VAL
        '    Cells(j + 1, 9) = P3_VAL
        '    Cells(j + 1, 10) = P4_VAL
        'End If
    Next j
    
' calculate average of all the FINAL portfolios
    MEAN1_VAL = MEAN1_VAL + P1_VAL
    MEAN2_VAL = MEAN2_VAL + P2_VAL
    MEAN3_VAL = MEAN3_VAL + P3_VAL
    MEAN4_VAL = MEAN4_VAL + P4_VAL
Next i

' calculate average of all the FINAL values
MEAN1_VAL = MEAN1_VAL / nLOOPS
MEAN2_VAL = MEAN2_VAL / nLOOPS
MEAN3_VAL = MEAN3_VAL / nLOOPS
MEAN4_VAL = MEAN4_VAL / nLOOPS

P_VAL = P_VAL / nLOOPS / PERIODS                ' average WIN_VAL percentage
WIN_VAL = WIN_VAL / nLOOPS / PERIODS            ' average WIN_VAL return
LOSS_VAL = LOSS_VAL / nLOOPS / PERIODS          ' average LOSS_VAL return
SD_WIN_VAL = Sqr(SD_WIN_VAL / nLOOPS / PERIODS)   ' standard deviation of wins
SD_LOSS_VAL = Sqr(SD_LOSS_VAL / nLOOPS / PERIODS) ' standard deviation of losses

ReDim TEMP_VECTOR(1 To 9, 1 To 2)

TEMP_VECTOR(1, 1) = "Actual Win prob"
TEMP_VECTOR(2, 1) = "Actual Win Return"
TEMP_VECTOR(3, 1) = "Actual Loss Return"
TEMP_VECTOR(4, 1) = "Actual Win Volatility"
TEMP_VECTOR(5, 1) = "Actual Loss Volatility"

TEMP_VECTOR(6, 1) = "Average K1 Portfolio (after " & _
                    Format(nLOOPS, "0") & " simulations)"

TEMP_VECTOR(7, 1) = "Average K2 Portfolio (after " & _
                    Format(nLOOPS, "0") & " simulations)"

TEMP_VECTOR(8, 1) = "Average K3 Portfolio (after " & _
                    Format(nLOOPS, "0") & " simulations)"

TEMP_VECTOR(9, 1) = "Average K4 Portfolio (after " & _
                    Format(nLOOPS, "0") & " simulations)"

TEMP_VECTOR(1, 2) = P_VAL
TEMP_VECTOR(2, 2) = WIN_VAL
TEMP_VECTOR(3, 2) = LOSS_VAL
TEMP_VECTOR(4, 2) = SD_WIN_VAL
TEMP_VECTOR(5, 2) = SD_LOSS_VAL
TEMP_VECTOR(6, 2) = MEAN1_VAL
TEMP_VECTOR(7, 2) = MEAN2_VAL
TEMP_VECTOR(8, 2) = MEAN3_VAL
TEMP_VECTOR(9, 2) = MEAN4_VAL

KELLY_RATIO_MC_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
KELLY_RATIO_MC_FUNC = Err.number
End Function

Function KELLY_RATIO_TRADING_SCHEME_FUNC(ByVal P_VAL As Double, _
ByVal W_VAL As Double, _
ByVal L_VAL As Double)

'Kelly Ratio Trading Scheme:
'K = P - (1-P)/{W/L}
'where
'P is the probability of winning (example: P = 0.5, meaning 50% of the time you expect to win)
'W is the expected dollar amount of your winnings per trade (example: W = $1.25)
'L is the expected dollar amount of your losses per trade (example: L = $1.00)
'and K is the fraction devoted to each trade (example: K = 0.5 - (1-0.5)/(1.25) = 0.1 or 10%.)

'Note that it's the ratio W/L that's important. You get the same result with W=$12,500 and L = $10,000.

On Error GoTo ERROR_LABEL

KELLY_RATIO_TRADING_SCHEME_FUNC = P_VAL - (1 - P_VAL) / (W_VAL / L_VAL)

Exit Function
ERROR_LABEL:
KELLY_RATIO_TRADING_SCHEME_FUNC = Err.number
End Function
