Attribute VB_Name = "FINAN_ASSET_SYSTEM_ATR_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'Over the years many commodity trading advisors, proprietary traders, and global macro hedge
'funds have successfully applied various trend following methods to profitably trade in global
'futures markets. Very little research, however, has been published regarding trend following
'strategies applied to stocks. Is it reasonable to assume that trend following works on futures
'but not stocks? We decided to put a long only trend following strategy to the test by running
'it against a comprehensive database of U.S. stocks that have been adjusted for corporate actions.
'Delisted companies were included to account for survivorship bias. Realistic transaction
'cost estimates (slippage & commission) were applied. Liquidity filters were used to limit
'hypothetical trading to only stocks that would have been liquid enough to trade, at the time of
'the trade. Coverage included 24,000+ securities spanning 22 years. The empirical results strongly
'suggest that trend following on stocks does offer a positive mathematical expectancy, an
'essential building block of an effective investing or trading system.

'References:
'http://gummy-stuff.org/trends2.htm
'http://www.blackstarfunds.com/files/Does_trendfollowing_work_on_stocks.pdf

Function ASSET_WEEKLY_ATR_SIGNAL_FUNC(ByRef TICKER_STR As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal END_DATE As Date, _
Optional ByVal ATR_AVG_WEEKS_PERIODS As Long = 8, _
Optional ByVal ATR_FACTOR As Double = 10, _
Optional ByVal MAX_WEEKS_PERIODS As Long = 2000, _
Optional ByVal BOLLI_WEEKS_PERIODS As Long = 40, _
Optional ByVal BOLLI_FACTOR As Double = 2.5, _
Optional ByVal INITIAL_CASH As Double = 100000, _
Optional ByVal INITIAL_SHARES As Long = 0, _
Optional ByVal VERSION As Integer = 0, _
Optional ByVal OUTPUT As Integer = 0)

'ATR_AVG_WEEKS_PERIODS: in periods
'ATR_FACTOR: ATR Multiplier

'MAX_WEEKS_PERIODS: Max weeks for a BUY

'You can choose either the ATR SELL signal or the Bolli?
'VERSION = 0 Then: ATR
'VERSION = 1 Then: Bolli
'VERSION => 2 Then: Trailling Stop

Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim NROWS As Long

Dim U_VAL As Double
Dim V_VAL As Double
Dim W_VAL As Double
Dim X_VAL As Double

Dim MAX_VAL As Double
Dim TEMP_VAL As Double

Dim TEMP1_SUM As Double
Dim TEMP2_SUM As Double
Dim TEMP3_SUM As Double
Dim TEMP4_SUM As Double

Dim MEAN_VAL As Double
Dim SIGMA_VAL As Double

Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

Const BUY_STR As String = "BUY"
Const SELL_STR As String = "SELL"

On Error GoTo ERROR_LABEL


'We download ten years worth of weekly stock prices.
'At the end of each week we look at the maximum of Friday's Closing stock Price, from ten
'years ago until this Friday's Close. we 're talking about an "all time high", from the
'starting week.

'If Friday's Close is at that Maximum, we assume the stock is on an uptrend ... and we BUY
'at Monday's Open.. At the end of each week we also calculate the True Range, averaged over
'the past 8 weeks. ... about 40 market days.

'The True Range (TR) identifies the largest price variation from last Friday's Close to this
'Friday's Close. We compare this Friday's Close to 10x the ATR value. If it's less than (or
'equal to) 10ATR, we SELL at the Open, on the following Monday.
'TR : The Maximum Variation in Price

'Sell when the Price falls below 10xATR? That seems curious. I mean ...
'Think of it this way:   ATR (averaged over the last 8 weeks), is a fraction f of the
'current Price P, namely f = ATR / P ... so ATR = f P.

'Example: f = ATR/P = 0.08, meaning the maximum price variation is 8% of the current Price.
'Then 10 ATR = 10 (f P) and that's when we Sell ... when the Price drops below that.

'Example: If f = ATR/P = 0.08, we sell when the Price drops below 10 f P = 0.80 P. That's a
'drop of 20%.

'For example, in the picture of the spreadsheet (below), there's a Sell signal on Friday,
'Oct 2, 1998. At that time, the Price was just $19.79 and ATR was $2.112.
'We Sell because the Price had dropped below 10 ATR = $21.12.
'We sell at Monday's Open (which happened to be $19.62.)
'The week before, P = $21.74 and 10 ATR = $17.30 ... so no Sell signal was generated.

'That 's it?
'Yes. We have a BUY signal, based upon the running maximum ... and a SELL signal based upon the
'Average True Range. Now we just ...

'Does it work good? That 's what I was going to talk about.

'After downloading data, we can pick the number of weeks to average (Example: 8 weeks) and
'the multiplier (Example: 10). We start with umpteen Shares and/or $Cash in our portfolio ...
'then follow the trend as described above. (intial Shares were 0 and initial Cash was $100K.

'But you get 12% Compound Annual Growth Rate instead of 11%? That's good?
'But it 'll depend upon the stock and the time period and what you choose for the number of weeks
'in the average and the ATR multiplier and the phases of the moon etc. etc.

'Note that, in this lousy market environment, you wind up selling everything and sitting on Cash.
'Of course, if the asset has a low volatility, you never sell. For example, here's a Bond Fund:
'So what about some other time period?

'You 'll (sometimes) find that, for certain time periods and parameter choices, you Sell and
'miss the subsequent rise in price 'cause it don't hardly reach the moving maximum ... after a
'big drop. So why don't you change the parameters, like that 10 or maybe the 8 ...

'If you had used the above scheme during the crash of 1929, you'd have sold ... then stayed
'in Cash for some 20 or 30 years. Suppose we only look back at the Max Price for, say, 50 weeks.

'Then we'd get something like this (for our earlier example of the GE):
'Now you won't miss those big increases after the big drop and ...
'So you really believe in this stuff?

'I have this bridge I'd like to sell you. It's in very good shape and ...
'I forgot to mention that there's a cell in the spreadsheet that says 2000.
'It means "calculate the Max over the past 2000 weeks" ... which really means
'from the initial time. If you leave it at 2000 you get the Buy/Sell ritual we'
've been explaining. BUT, it you change it to, say, 50, you BUY when the price is
'at the Max over the past 50 weeks.

'Since we 've modified the BUY signal (so we only look at the Max over the past umpteen
'weeks rather than over all weeks), we should ... Modify the SELL signal as well.
'You took the words right outta my mouth.

'I reckon there are lots of ways to do that, but let's so something familar: the Lower
'Bollinger Band. Recall that we calculate the Average Price over the past umpteen weeks as
'well as the Standard Deviation of Prices over this same time period.

'Then the Lower Bolli is: (Average Price) - k (Standard Deviation) where k is some number ...
'like maybe 2 or 3. So, when the Price drops to 2 Standard Deviations below the Moving Average
'Price, we figure it's dropped enough ... so we SELL.

If IsArray(TICKER_STR) = False Then
    DATA_MATRIX = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, _
                  START_DATE, END_DATE, "w", "DOHLCVA", False, _
                  True, True)
Else
    DATA_MATRIX = TICKER_STR
End If
If IsArray(DATA_MATRIX) = False Then: GoTo ERROR_LABEL

NROWS = UBound(DATA_MATRIX, 1)

If VERSION = 0 Then
    X_VAL = 0
ElseIf VERSION = 1 Then
    X_VAL = 1
Else 'If VERSION >= 2 Then
    X_VAL = 2
End If

U_VAL = (X_VAL - 1) * (X_VAL - 2) / 2
'=(X-1)*(X-2)/2

V_VAL = X_VAL * (2 - X_VAL)
'=X*(2-X)

W_VAL = X_VAL * (X_VAL - 1) / 2
'=X*(X-1)/2

'-----------------------------------------------------------------------------
ReDim TEMP_MATRIX(0 To NROWS, 1 To 31)
'-----------------------------------------------------------------------------
TEMP_MATRIX(0, 1) = "DATE"
TEMP_MATRIX(0, 2) = "OPEN"
TEMP_MATRIX(0, 3) = "HIGH"
TEMP_MATRIX(0, 4) = "LOW"
TEMP_MATRIX(0, 5) = "CLOSE"
TEMP_MATRIX(0, 6) = "VOLUME"
TEMP_MATRIX(0, 7) = "ADJ CLOSE"
TEMP_MATRIX(0, 8) = "RETURNS"
'-----------------------------------------------------------------------------
TEMP_MATRIX(0, 9) = "H-L"
TEMP_MATRIX(0, 10) = "H-pC"
TEMP_MATRIX(0, 11) = "L-pC"
TEMP_MATRIX(0, 12) = "TR"
'-----------------------------------------------------------------------------
TEMP_MATRIX(0, 13) = Format(ATR_FACTOR, "0.0") & "x" & ATR_AVG_WEEKS_PERIODS & "ATR"
TEMP_MATRIX(0, 14) = "MAX PRICE"
TEMP_MATRIX(0, 15) = "SCALE"
TEMP_MATRIX(0, 16) = "MONDAY TRADES"
'-----------------------------------------------------------------------------
TEMP_MATRIX(0, 17) = "SHARES"
TEMP_MATRIX(0, 18) = "CASH"
TEMP_MATRIX(0, 19) = "PORTFOLIO"
'-----------------------------------------------------------------------------
TEMP_MATRIX(0, 20) = "SCALED_OPEN"
TEMP_MATRIX(0, 21) = "SCALED_HIGH"
TEMP_MATRIX(0, 22) = "SCALED_LOW"
TEMP_MATRIX(0, 23) = "SCALED_CLOSE"
'-----------------------------------------------------------------------------
TEMP_MATRIX(0, 24) = "BUY & HOLD"
TEMP_MATRIX(0, 25) = "BUY MARKET"
TEMP_MATRIX(0, 26) = "SELL MARKET"
TEMP_MATRIX(0, 27) = "FIRST TIME BUY"
TEMP_MATRIX(0, 28) = "FIRST TIME SELL"
'-----------------------------------------------------------------------------
TEMP_MATRIX(0, 29) = "STDEV"
TEMP_MATRIX(0, 30) = "LOWER BOLLI"
TEMP_MATRIX(0, 31) = "TRAILING STOP"
'-----------------------------------------------------------------------------

For j = 1 To 7: TEMP_MATRIX(1, j) = DATA_MATRIX(1, j): Next j
TEMP_MATRIX(1, 6) = TEMP_MATRIX(1, 6) / 1000
    
TEMP_MATRIX(1, 15) = DATA_MATRIX(1, 7) / DATA_MATRIX(1, 5)
    
TEMP_MATRIX(1, 20) = TEMP_MATRIX(1, 2) * TEMP_MATRIX(1, 15)
TEMP_MATRIX(1, 21) = TEMP_MATRIX(1, 3) * TEMP_MATRIX(1, 15)
TEMP_MATRIX(1, 22) = TEMP_MATRIX(1, 4) * TEMP_MATRIX(1, 15)
TEMP_MATRIX(1, 23) = TEMP_MATRIX(1, 5) * TEMP_MATRIX(1, 15)

TEMP1_SUM = 0
TEMP2_SUM = 0
        
TEMP_MATRIX(1, 8) = DATA_MATRIX(1, 5) / DATA_MATRIX(1, 2) - 1
TEMP_MATRIX(1, 13) = 0
TEMP_MATRIX(1, 14) = TEMP_MATRIX(1, 23)
TEMP_MATRIX(1, 15) = DATA_MATRIX(1, 7) / DATA_MATRIX(1, 5)
TEMP_MATRIX(1, 17) = INITIAL_SHARES
TEMP_MATRIX(1, 18) = INITIAL_CASH
TEMP_MATRIX(1, 19) = TEMP_MATRIX(1, 18) + TEMP_MATRIX(1, 17) * TEMP_MATRIX(1, 23)
TEMP_MATRIX(1, 24) = TEMP_MATRIX(1, 19)

TEMP_MATRIX(1, 30) = 0
TEMP_MATRIX(1, 31) = 0

l = 1
MAX_VAL = TEMP_MATRIX(1, 14)
TEMP2_SUM = TEMP_MATRIX(1, 23)
TEMP4_SUM = 0

'-----------------------------------------------------------------------------
For i = 2 To NROWS
'-----------------------------------------------------------------------------
    For j = 1 To 7: TEMP_MATRIX(i, j) = DATA_MATRIX(i, j): Next j
    TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 6) / 1000
    
    TEMP_MATRIX(i, 15) = DATA_MATRIX(i, 7) / DATA_MATRIX(i, 5)
    
    TEMP_MATRIX(i, 20) = TEMP_MATRIX(i, 2) * TEMP_MATRIX(i, 15)
    TEMP_MATRIX(i, 21) = TEMP_MATRIX(i, 3) * TEMP_MATRIX(i, 15)
    TEMP_MATRIX(i, 22) = TEMP_MATRIX(i, 4) * TEMP_MATRIX(i, 15)
    TEMP_MATRIX(i, 23) = TEMP_MATRIX(i, 5) * TEMP_MATRIX(i, 15)
    
    TEMP_MATRIX(i, 8) = DATA_MATRIX(i, 7) / DATA_MATRIX(i - 1, 7) - 1
        
    TEMP_MATRIX(i, 9) = TEMP_MATRIX(i, 21) - TEMP_MATRIX(i, 22)
    TEMP_VAL = TEMP_MATRIX(i, 9)
        
    TEMP_MATRIX(i, 10) = Abs(TEMP_MATRIX(i, 21) - TEMP_MATRIX(i - 1, 23))
    If TEMP_MATRIX(i, 10) > TEMP_VAL Then: TEMP_VAL = TEMP_MATRIX(i, 10)
        
    TEMP_MATRIX(i, 11) = Abs(TEMP_MATRIX(i, 22) - TEMP_MATRIX(i - 1, 23))
    If TEMP_MATRIX(i, 11) > TEMP_VAL Then: TEMP_VAL = TEMP_MATRIX(i, 11)
        
    TEMP_MATRIX(i, 12) = TEMP_VAL
        
    If i <= ATR_AVG_WEEKS_PERIODS + 1 Then
        TEMP1_SUM = TEMP1_SUM + TEMP_MATRIX(i, 12)
        k = i - 1
        TEMP_MATRIX(i, 13) = ATR_FACTOR * TEMP1_SUM / k
    Else
        k = i - ATR_AVG_WEEKS_PERIODS - 1
        TEMP1_SUM = TEMP1_SUM - TEMP_MATRIX(k, 12)
        TEMP1_SUM = TEMP1_SUM + TEMP_MATRIX(i, 12)
            
        k = ATR_AVG_WEEKS_PERIODS + 1
        TEMP_MATRIX(i, 13) = ATR_FACTOR * TEMP1_SUM / k
    End If
    
    l = l + 1
    If l > MAX_WEEKS_PERIODS Then
        MAX_VAL = TEMP_MATRIX(i, 23)
        l = 1
    End If
        
    If TEMP_MATRIX(i, 23) > MAX_VAL Then: MAX_VAL = TEMP_MATRIX(i, 23)
    TEMP_MATRIX(i, 14) = MAX_VAL
        
    If i <= BOLLI_WEEKS_PERIODS + 1 Then
        TEMP2_SUM = TEMP2_SUM + TEMP_MATRIX(i, 23)
        MEAN_VAL = TEMP2_SUM / i
            
        TEMP3_SUM = 0
        For j = 1 To i
            TEMP3_SUM = TEMP3_SUM + (TEMP_MATRIX(j, 23) - MEAN_VAL) ^ 2
        Next j
        
        TEMP_MATRIX(i, 29) = (TEMP3_SUM / i) ^ 0.5
        TEMP_MATRIX(i, 30) = MEAN_VAL - (BOLLI_FACTOR * TEMP_MATRIX(i, 29))
    Else
    
        k = i - BOLLI_WEEKS_PERIODS - 1
        
        TEMP2_SUM = TEMP2_SUM - TEMP_MATRIX(k, 23)
        TEMP2_SUM = TEMP2_SUM + TEMP_MATRIX(i, 23)
        
        MEAN_VAL = TEMP2_SUM / (BOLLI_WEEKS_PERIODS + 1)
        
        k = k + 1
        TEMP3_SUM = 0
        For j = i To k Step -1
            TEMP3_SUM = TEMP3_SUM + (TEMP_MATRIX(j, 23) - MEAN_VAL) ^ 2
        Next j
        
        TEMP_MATRIX(i, 29) = (TEMP3_SUM / (BOLLI_WEEKS_PERIODS + 1)) ^ 0.5
        TEMP_MATRIX(i, 30) = MEAN_VAL - (BOLLI_FACTOR * TEMP_MATRIX(i, 29))

    End If
    
    TEMP_MATRIX(i, 24) = TEMP_MATRIX(i - 1, 24) * TEMP_MATRIX(i, 23) / TEMP_MATRIX(i - 1, 23)


    If CDbl(TEMP_MATRIX(i - 1, 23)) >= CDbl(TEMP_MATRIX(i - 1, 14)) Then
        TEMP_MATRIX(i, 16) = BUY_STR
    Else
        If CDbl(TEMP_MATRIX(i - 1, 23)) <= CDbl(U_VAL * TEMP_MATRIX(i - 1, 13) + _
                                     V_VAL * TEMP_MATRIX(i - 1, 30) + _
                                     W_VAL * TEMP_MATRIX(i - 1, 31)) Then
            TEMP_MATRIX(i, 16) = SELL_STR
        Else
            TEMP_MATRIX(i, 16) = ""
        End If
    End If
    
    
    If TEMP_MATRIX(i, 16) = SELL_STR Then
        TEMP_MATRIX(i, 17) = 0
    Else
        If TEMP_MATRIX(i, 16) = BUY_STR Then
            TEMP_MATRIX(i, 17) = TEMP_MATRIX(i - 1, 17) + _
                                 TEMP_MATRIX(i - 1, 18) / TEMP_MATRIX(i, 20)
        Else
            TEMP_MATRIX(i, 17) = TEMP_MATRIX(i - 1, 17)
        End If
    End If
    
    
    If TEMP_MATRIX(i, 16) = SELL_STR Then
        TEMP_MATRIX(i, 18) = TEMP_MATRIX(i - 1, 18) + _
                             TEMP_MATRIX(i, 20) * TEMP_MATRIX(i - 1, 17)
    Else
        If TEMP_MATRIX(i, 16) = BUY_STR Then
            TEMP_MATRIX(i, 18) = 0
        Else
            TEMP_MATRIX(i, 18) = TEMP_MATRIX(i - 1, 18)
        End If
    End If
    
    TEMP_MATRIX(i, 19) = TEMP_MATRIX(i, 18) + TEMP_MATRIX(i, 17) * TEMP_MATRIX(i, 23)
    TEMP4_SUM = TEMP4_SUM + (TEMP_MATRIX(i, 19) / TEMP_MATRIX(i - 1, 19) - 1)

    If i <> 2 Then
        If (TEMP_MATRIX(i, 16) = BUY_STR And TEMP_MATRIX(i, 13) > TEMP_MATRIX(i - 1, 31)) Then
            TEMP_MATRIX(i, 31) = TEMP_MATRIX(i, 13)
        Else
            TEMP_MATRIX(i, 31) = TEMP_MATRIX(i - 1, 31)
        End If
    Else
        TEMP_MATRIX(i, 31) = TEMP_MATRIX(i, 13)
    End If
    
    If (TEMP_MATRIX(i, 16) = SELL_STR And TEMP_MATRIX(i - 1, 18) = 0) Then
        TEMP_MATRIX(i, 28) = 1
    Else
        TEMP_MATRIX(i, 28) = 0
    End If
    
    If ((TEMP_MATRIX(i, 16) = BUY_STR) And (TEMP_MATRIX(i - 1, 17) = 0)) Then
        TEMP_MATRIX(i, 27) = 1
    Else
        TEMP_MATRIX(i, 27) = 0
    End If
    
    If TEMP_MATRIX(i, 16) = SELL_STR Then
        TEMP_MATRIX(i, 26) = TEMP_MATRIX(i, 28) * TEMP_MATRIX(i, 24)
    Else
        TEMP_MATRIX(i, 26) = 0
    End If
    
    If TEMP_MATRIX(i, 16) = BUY_STR Then
        TEMP_MATRIX(i, 25) = TEMP_MATRIX(i, 27) * TEMP_MATRIX(i, 24)
    Else
        TEMP_MATRIX(i, 25) = 0
    End If
'-----------------------------------------------------------------------------
Next i
'-----------------------------------------------------------------------------

Select Case OUTPUT
Case 0
    ASSET_WEEKLY_ATR_SIGNAL_FUNC = TEMP_MATRIX
Case Else
    MEAN_VAL = TEMP4_SUM / (NROWS - 1)
    SIGMA_VAL = 0
    For i = 2 To NROWS
        SIGMA_VAL = SIGMA_VAL + ((TEMP_MATRIX(i, 19) / TEMP_MATRIX(i - 1, 19) - 1) - MEAN_VAL) ^ 2
    Next i
    SIGMA_VAL = (SIGMA_VAL / (NROWS - 1)) ^ 0.5
    If OUTPUT = 1 Then
        ASSET_WEEKLY_ATR_SIGNAL_FUNC = MEAN_VAL / SIGMA_VAL
    Else
        ASSET_WEEKLY_ATR_SIGNAL_FUNC = Array(MEAN_VAL / SIGMA_VAL, MEAN_VAL, SIGMA_VAL)
    End If
End Select

Exit Function
ERROR_LABEL:
ASSET_WEEKLY_ATR_SIGNAL_FUNC = "--"
End Function


'----------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------

'ATR Calculation for purpuse of setting hard and trailing stops on LONG positions only
'Idea is that rather than using a 5% position, with a 4% stop (risking 0.2% of portfolio per
'trade), we determine the portfolio weight based on the hard strop assuming X-factor of ATR
'Another option would be to base position size on units and pyrimid into positions, maybe based
'on X-ATRs above the purchase price to add sequential units (max of say 4 units like the turtles)
'Position Size = (Account Liquidation Value * Risk per Trade %) / (Hard Stop Multiplier * ATR)

Function ASSETS_ATR_HARD_TRAILING_STOPS_FUNC(ByVal TICKERS_RNG As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal END_DATE As Date, _
Optional ByVal ATR_PERIOD As Long = 20, _
Optional ByVal ATR_HARD_STOP_MULT As Double = 1.5, _
Optional ByVal ATR_TRAILING_STOP_MULT As Double = 4, _
Optional ByVal ACCOUNT_BALANCE As Double = 1000000, _
Optional ByVal RISK_PER_TRADE As Double = 0.002)

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim NSIZE As Long

Dim A0_VAL As Double
Dim A1_VAL As Double

Dim TEMP_VAL As Double
Dim MULT_VAL As Double 'ATR_MULT

Dim LOW_VAL As Double
Dim HIGH_VAL As Double

Dim MAX_VAL As Double

Dim TICKER_STR As String
Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant
Dim TICKERS_VECTOR As Variant

On Error GoTo ERROR_LABEL

If IsArray(TICKERS_RNG) Then
    TICKERS_VECTOR = TICKERS_RNG
Else
    ReDim TICKERS_VECTOR(1 To 1, 1 To 1)
    TICKERS_VECTOR(1, 1) = TICKERS_RNG
End If
NSIZE = UBound(TICKERS_VECTOR, 1)
MULT_VAL = 2 / (ATR_PERIOD + 1)

ReDim TEMP_MATRIX(0 To NSIZE, 1 To 9)
TEMP_MATRIX(0, 1) = "SYMBOL"
TEMP_MATRIX(0, 2) = "PRIOR DAY CLOSING PRICE"
TEMP_MATRIX(0, 3) = ATR_PERIOD & "-DAY ATR VALUE"
TEMP_MATRIX(0, 4) = "HARD STOP AMOUNT"
TEMP_MATRIX(0, 5) = "HARD STOP %"
TEMP_MATRIX(0, 6) = "TRAILING STOP AMOUNT"
TEMP_MATRIX(0, 7) = "TRAILING STOP %"
TEMP_MATRIX(0, 8) = "POSITION SIZE (# OF SHARES)"
TEMP_MATRIX(0, 9) = "PORTFOLIO WEIGHT %"

For j = 1 To NSIZE
    TICKER_STR = TICKERS_VECTOR(j, 1)
    TEMP_MATRIX(j, 1) = TICKER_STR
    DATA_MATRIX = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, START_DATE, END_DATE, _
                  "d", "DOHLCVA", False, True, True)
    If IsArray(DATA_MATRIX) = False Then: GoTo 1983
    NROWS = UBound(DATA_MATRIX, 1)

    A0_VAL = 0: A1_VAL = 0 'ATR values for today and yesterday

    i = 1 'day one
    MAX_VAL = DATA_MATRIX(i, 3) - DATA_MATRIX(i, 4) 'calculates TR on day one
    A0_VAL = MAX_VAL 'calculates ATR on day one for next day
    
    For i = 2 To NROWS
        TEMP_VAL = DATA_MATRIX(i, 3) - DATA_MATRIX(i, 4) 'calculates intraday TR
        HIGH_VAL = DATA_MATRIX(i, 3) - DATA_MATRIX(i - 1, 5)
        If HIGH_VAL < 0 Then: HIGH_VAL = HIGH_VAL * -1
        LOW_VAL = DATA_MATRIX(i, 4) - DATA_MATRIX(i - 1, 5)
        If LOW_VAL < 0 Then: LOW_VAL = LOW_VAL * -1
        
        If TEMP_VAL >= HIGH_VAL And TEMP_VAL >= LOW_VAL Then 'determines the max TR
            MAX_VAL = TEMP_VAL
        ElseIf HIGH_VAL >= TEMP_VAL And HIGH_VAL >= LOW_VAL Then
            MAX_VAL = HIGH_VAL
        ElseIf LOW_VAL >= TEMP_VAL And LOW_VAL >= HIGH_VAL Then
            MAX_VAL = LOW_VAL
        End If
    
        A1_VAL = (MAX_VAL - A0_VAL) * MULT_VAL + A0_VAL
        A0_VAL = A1_VAL
        
    Next i
    TEMP_MATRIX(j, 2) = DATA_MATRIX(NROWS, 5) 'Prior Day Closing Price
    TEMP_MATRIX(j, 3) = A1_VAL
    TEMP_MATRIX(j, 4) = A1_VAL * ATR_HARD_STOP_MULT 'Calculates hard stop amount
    TEMP_MATRIX(j, 5) = TEMP_MATRIX(j, 4) / DATA_MATRIX(NROWS, 5)
    TEMP_MATRIX(j, 6) = A1_VAL * ATR_TRAILING_STOP_MULT 'Calculates trailing stop amount
    TEMP_MATRIX(j, 7) = TEMP_MATRIX(j, 6) / DATA_MATRIX(NROWS, 5)
    
    If TEMP_MATRIX(j, 4) > 0 Then
        TEMP_MATRIX(j, 8) = (ACCOUNT_BALANCE * RISK_PER_TRADE) / TEMP_MATRIX(j, 4)
        'Risk per trade, based on hard stop
    Else
        TEMP_MATRIX(j, 8) = 1
    End If
    TEMP_MATRIX(j, 9) = (TEMP_MATRIX(j, 2) * TEMP_MATRIX(j, 8)) / ACCOUNT_BALANCE
1983:
Next j

ASSETS_ATR_HARD_TRAILING_STOPS_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
ASSETS_ATR_HARD_TRAILING_STOPS_FUNC = Err.number
End Function
