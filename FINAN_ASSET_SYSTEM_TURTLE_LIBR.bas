Attribute VB_Name = "FINAN_ASSET_SYSTEM_TURTLE_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'References:
'
'http://www.financialwebring.org/gummystuff/turtle-trading.htm
'http://www.financialwebring.org/gummystuff/ATR.htm
'http://www.financialwebring.org/gummystuff/Bollinger.htm#EMA

Function ASSET_TURTLE_SIGNAL_FUNC(ByRef TICKER_STR As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal END_DATE As Date, _
Optional ByVal DONCHIAN_PERIOD As Double = 20, _
Optional ByVal EMA_PERIOD As Double = 20, _
Optional ByVal INITIAL_CASH As Double = 100000, _
Optional ByVal TURTLE_RULE As Double = 0.01, _
Optional ByVal COUNT_BASIS As Double = 252, _
Optional ByVal VERSION As Integer = 1, _
Optional ByVal OUTPUT As Integer = 0)

'IIf(VERSION = 0, "ATR", "APR")

'-------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------
'UNIT gives the number of Shares your should trade?
'Aah, very good question, however, lets' compare the logic behind the two schemes,
'assuming that (1% of Equity) = $100K
'-------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------
'UNIT = (1% of Equity)/[ (ATR)*(Stock Price) ]
'* If Stock Price = $50, then (1% of Equity)/(Stock Price) = 2,000
'... the maximum number of shares that $100K will buy.
'* Instead we buy UNIT shares.
'* Suppose the average/maximum 20-day variation for one share is ATR = $1.25.
'* Then the variation for UNIT shares is UNIT*ATR =UNIT*1.25.
'* We set UNIT*ATR = (1% of Equity)/(Stock Price) = 2,000.
'* Then UNIT = 2000/1.25 = 1600 shares. (Cost = $50*1600 = $80K.)
'That defines the ATR UNIT.
    

'UNIT = (1% of Equity)/[ (APR)*(Stock Price) ]
'* If Stock Price = $50, then (1% of Equity)/(Stock Price) = 2,000
'... the maximum number of shares that $100K will buy.
'* Instead we buy UNIT shares.
'* Suppose the average/maximum 20-day percentage variation for one share is APR = 2.5%.
'* Then the percentage variation for UNIT shares is UNIT*APR = UNIT*0.025.
'* We set UNIT*APR = (1% of Equity)/(Stock Price) = 2,000.
'* Then UNIT = 2000/0.025 = 80,000 shares. (Cost = almost infinite!)
'That defines the APR UNIT.

'Cost = almost infinite! You kidding?
'Yes , it 's true. Dividing by APR, a percentage, we'd get a HUGE value for our UNIT.
'But we 'll fix that by dividing by 100*APR. Then, if APR = 2.5%, we'll just divide by 2.5.
'In the above example, the APR UNIT would then by 2000/2.5 = 800 shares at a cost of $50*800 = $40K.
'-------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long
Dim m As Long

Dim NO_DAYS As Long
Dim NROWS As Long
Dim NCOLUMNS As Long

Dim MAX_VAL As Double
Dim MIN_VAL As Double
Dim TEMP_SUM As Double

Dim SCAGR_VAL As Double
Dim XCAGR_VAL As Double

Dim COST_VAL As Double
Dim ALPHA_VAL As Double

Dim MEAN_VAL As Double
Dim SIGMA_VAL As Double

Dim DATA_MATRIX As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

ALPHA_VAL = 1 - 1 / EMA_PERIOD

If IsArray(TICKER_STR) = True Then
    DATA_MATRIX = TICKER_STR
Else
    DATA_MATRIX = _
    YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, START_DATE, END_DATE, _
    "d", "DOHLCVA", False, False, True)
End If
NROWS = UBound(DATA_MATRIX, 1)

'--------------------------------------------------------------------------------
m = 2
NCOLUMNS = 21
ReDim TEMP_MATRIX(0 To NROWS, 1 To NCOLUMNS)
For i = 0 To NROWS: For j = 1 To NCOLUMNS: TEMP_MATRIX(i, j) = "": Next j: Next i
'--------------------------------------------------------------------------------
TEMP_MATRIX(0, 1) = "DATE"
TEMP_MATRIX(0, 2) = "OPEN"
TEMP_MATRIX(0, 3) = "HIGH"
TEMP_MATRIX(0, 4) = "LOW"
TEMP_MATRIX(0, 5) = "CLOSE"
TEMP_MATRIX(0, 6) = "VOLUME"

TEMP_MATRIX(0, 7) = "ADJ CLOSE" '7
TEMP_MATRIX(0, 8) = "RETURN" '8

TEMP_MATRIX(0, 9) = "HIGH - LOW" '9
TEMP_MATRIX(0, 10) = "HIGH - pCLOSE" '10
TEMP_MATRIX(0, 11) = "LOW - pCLOSE" '11

TEMP_MATRIX(0, 14) = "DON - HIGH" '14
TEMP_MATRIX(0, 15) = "DON - LOW" '15

TEMP_MATRIX(0, 17) = "SHARES TRADED" '17
TEMP_MATRIX(0, 18) = "SHARE$ @ OPEN" '18
TEMP_MATRIX(0, 19) = "SHARES HELD" '19
TEMP_MATRIX(0, 20) = "CASH" '20

COST_VAL = INITIAL_CASH
If VERSION = 0 Then 'ATR --> Dollar Value
    COST_VAL = COST_VAL * TURTLE_RULE / 1
Else 'APR --> Percent
    COST_VAL = COST_VAL * TURTLE_RULE / 100
End If

'--------------------------------------------------------------------------------
i = 1
For j = 1 To 7: TEMP_MATRIX(i, j) = DATA_MATRIX(i, j): Next j
TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 6) / 1000
For j = 8 To 15: TEMP_MATRIX(i, j) = "": Next j
TEMP_MATRIX(i, 13) = 0
'--------------------------------------------------------------------------------

h = 0: TEMP_SUM = 0
For i = 2 To NROWS
    For j = 1 To 7: TEMP_MATRIX(i, j) = DATA_MATRIX(i, j): Next j
    TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 6) / 1000
    TEMP_MATRIX(i, 8) = DATA_MATRIX(i, 7) / DATA_MATRIX(i - 1, 7) - 1
        
    If VERSION = 0 Then 'ATR
        TEMP_MATRIX(i, 9) = TEMP_MATRIX(i, 3) - TEMP_MATRIX(i, 4)
        TEMP_MATRIX(i, 10) = Abs(TEMP_MATRIX(i, 3) - TEMP_MATRIX(i - 1, 5))
        TEMP_MATRIX(i, 11) = Abs(TEMP_MATRIX(i, 4) - TEMP_MATRIX(i - 1, 5))
    Else 'APR
        TEMP_MATRIX(i, 9) = (TEMP_MATRIX(i, 3) - TEMP_MATRIX(i, 4)) / TEMP_MATRIX(i - 1, 5)
        TEMP_MATRIX(i, 10) = Abs(TEMP_MATRIX(i, 3) / TEMP_MATRIX(i - 1, 5) - 1)
        TEMP_MATRIX(i, 11) = Abs(TEMP_MATRIX(i, 4) / TEMP_MATRIX(i - 1, 5) - 1)
    End If
    MAX_VAL = TEMP_MATRIX(i, 9)
    If TEMP_MATRIX(i, 10) > MAX_VAL Then: MAX_VAL = TEMP_MATRIX(i, 10)
    If TEMP_MATRIX(i, 11) > MAX_VAL Then: MAX_VAL = TEMP_MATRIX(i, 11)
    TEMP_MATRIX(i, 12) = MAX_VAL
    
    TEMP_MATRIX(i, 13) = ALPHA_VAL * TEMP_MATRIX(i - 1, 13) + (1 - ALPHA_VAL) * TEMP_MATRIX(i, 12)
        
    MAX_VAL = TEMP_MATRIX(i, 3)
    MIN_VAL = TEMP_MATRIX(i, 4)
    
    If i <= DONCHIAN_PERIOD Then k = 1 Else k = i - DONCHIAN_PERIOD
    For j = i To k Step -1
        If TEMP_MATRIX(j, 3) > MAX_VAL Then: MAX_VAL = TEMP_MATRIX(j, 3)
        If TEMP_MATRIX(j, 4) < MIN_VAL Then: MIN_VAL = TEMP_MATRIX(j, 4)
    Next j
        
    TEMP_MATRIX(i, 14) = MAX_VAL
    TEMP_MATRIX(i, 15) = MIN_VAL
    
    '-----------------------------------------------------------------------------------
    If i = EMA_PERIOD + DONCHIAN_PERIOD + 2 Then
    '-----------------------------------------------------------------------------------
        TEMP_MATRIX(i, 16) = Round(COST_VAL / TEMP_MATRIX(i, 13) / TEMP_MATRIX(i, 5), 0)
        TEMP_MATRIX(i, 17) = 0
        TEMP_MATRIX(i, 18) = TEMP_MATRIX(i, 17) * TEMP_MATRIX(i, 2) 'Open --> Assuming at Open
        TEMP_MATRIX(i, 19) = 0
        TEMP_MATRIX(i, 20) = INITIAL_CASH
        TEMP_MATRIX(i, 21) = TEMP_MATRIX(i, 20) + TEMP_MATRIX(i, 19) * TEMP_MATRIX(i, 5)
        l = i + 1
    '-----------------------------------------------------------------------------------
    ElseIf i > EMA_PERIOD + DONCHIAN_PERIOD + 2 Then
    '-----------------------------------------------------------------------------------

        TEMP_MATRIX(i, 16) = Round(COST_VAL / TEMP_MATRIX(i, 13) / TEMP_MATRIX(i, 5), 0)
                
        If (TEMP_MATRIX(i, 2) > TEMP_MATRIX(i - 1, 14) And TEMP_MATRIX(i - 1, 20) > TEMP_MATRIX(i, 16) * TEMP_MATRIX(i, 2)) Then
                TEMP_MATRIX(i, 17) = TEMP_MATRIX(i, 16) * 1
        Else
            If (TEMP_MATRIX(i, 2) < TEMP_MATRIX(i - 1, 15) And TEMP_MATRIX(i - 1, 19) > TEMP_MATRIX(i, 16)) Then
                TEMP_MATRIX(i, 17) = TEMP_MATRIX(i, 16) * -1
            Else
                TEMP_MATRIX(i, 17) = TEMP_MATRIX(i, 16) * 0
            End If
        End If
        
        TEMP_MATRIX(i, 18) = TEMP_MATRIX(i, 17) * TEMP_MATRIX(i, 2)
        
        TEMP_MATRIX(i, 19) = TEMP_MATRIX(i - 1, 19) + TEMP_MATRIX(i, 17)
        
        TEMP_MATRIX(i, 20) = TEMP_MATRIX(i - 1, 20) - TEMP_MATRIX(i, 18)
        
        TEMP_MATRIX(i, 21) = TEMP_MATRIX(i, 20) + TEMP_MATRIX(i, 19) * TEMP_MATRIX(i, 5)
        
        TEMP_SUM = TEMP_SUM + (TEMP_MATRIX(i, 21) / TEMP_MATRIX(i - 1, 21) - 1)
        h = h + 1
    '-----------------------------------------------------------------------------------
    End If
    '-----------------------------------------------------------------------------------
Next i

NO_DAYS = TEMP_MATRIX(NROWS, 1) - TEMP_MATRIX(1, 1)

XCAGR_VAL = (TEMP_MATRIX(NROWS, 5) / TEMP_MATRIX(l - 1, 5)) ^ (COUNT_BASIS / NO_DAYS) - 1
SCAGR_VAL = (TEMP_MATRIX(NROWS, 21) / TEMP_MATRIX(l - 1, 21)) ^ (COUNT_BASIS / NO_DAYS) - 1

TEMP_MATRIX(0, 21) = "SYSTEM @ CLOSE: CAGR = " & Format(SCAGR_VAL, "0.0000%")

If VERSION = 0 Then 'ATR
    TEMP_MATRIX(0, 12) = "ATR"
    TEMP_MATRIX(0, 13) = EMA_PERIOD & "-day ATR = " & Format(TEMP_MATRIX(NROWS, 13), "0.00")
    'In Dollar Value
Else 'APR
    TEMP_MATRIX(0, 12) = "APR"
    TEMP_MATRIX(0, 13) = EMA_PERIOD & "-day APR = " & Format(TEMP_MATRIX(NROWS, 13), "0.00%")
    'In Percent Value
End If
TEMP_MATRIX(0, 16) = "GST: " & Round((COST_VAL / TEMP_MATRIX(NROWS, 5)) / TEMP_MATRIX(NROWS, 13), 0)
'-----------------------------------------------------------------------------------
Select Case OUTPUT
'-----------------------------------------------------------------------------------
Case 0
'-----------------------------------------------------------------------------------
    ASSET_TURTLE_SIGNAL_FUNC = TEMP_MATRIX
'-----------------------------------------------------------------------------------
Case 1
'-----------------------------------------------------------------------------------
    ASSET_TURTLE_SIGNAL_FUNC = Array(SCAGR_VAL, XCAGR_VAL)
'-----------------------------------------------------------------------------------
Case Else
'-----------------------------------------------------------------------------------
    If h = 0 Then: GoTo ERROR_LABEL
    MEAN_VAL = TEMP_SUM / h
    SIGMA_VAL = 0
    For i = l To NROWS
        SIGMA_VAL = SIGMA_VAL + ((TEMP_MATRIX(i, 21) / TEMP_MATRIX(i - 1, 21) - 1) - MEAN_VAL) ^ 2
    Next i
    SIGMA_VAL = (SIGMA_VAL / h) ^ 0.5
    If OUTPUT = 2 Then
        ASSET_TURTLE_SIGNAL_FUNC = MEAN_VAL / SIGMA_VAL
    Else
        ASSET_TURTLE_SIGNAL_FUNC = Array(MEAN_VAL / SIGMA_VAL, MEAN_VAL, SIGMA_VAL)
    End If
'-----------------------------------------------------------------------------------
End Select
'-----------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
ASSET_TURTLE_SIGNAL_FUNC = "--"
End Function


Function ASSETS_TURTLE_TRADING_GST_FUNC(ByRef TICKERS_RNG As Variant, _
ByVal START_DATE As Date, _
ByVal END_DATE As Date, _
Optional ByVal EMA_PERIOD As Long = 20, _
Optional ByVal INITIAL_CASH As Double = 100000, _
Optional ByVal TURTLE_RULE As Double = 0.01)

'You type in up to X stock symbols and you get a set of associated gSTs and the COST of buying
'that many shares at the current stock price.

'Look closely and you can see that you should spend twice as much money on AA as AIG: $10K vs $5K.
'It's quite a different ratio if you use ATR.
'Indeed, their APRs are 10% and 20% so you see that the more volatile stock ...

'So how do things compare when you use ATR instead of APR?
'Aaah, good question! We consider the two prescriptions:

'For ATR we use the maximum value of:
'* (today's High) - (today's Low)
'* | (today's High) - (yesterday's Close) |
'* | (today's Low) - (yesterday's Close) |

'For APR we use the maximum value of:

'* 100 [ (today's High)- (today's Low) ] / (yesterday's Close)
'* 100 | (today's High) / (yesterday's Close) - 1 |
'* 100 | (today's Low) / (yesterday's Close) - 1|

'In other words, for APR, the ratios are relative to (yesterday's Close).
'That makes them percentages. That makes them independent of the currency. In Japanese yen
'or British pounds we'd get the same ...

'Are you going to answer my question?
'Huh? Oh, yes ... how they compare. Check out Table 2:
'Uh ... so how much of each should we you buy? Can you show the relative amounts so that ...?
'Okay, suppose we compare each investment with a $100 investment in that first stock: AA.

Dim i As Long
Dim j As Long
Dim NROWS As Long
Dim NCOLUMNS As Long

Dim ALPHA_VAL As Double

Dim ATR_VAL As Double
Dim ATR_HILO_VAL As Double
Dim ATR_HIPC_VAL As Double
Dim ATR_LOPC_VAL As Double
Dim ATR_FACTOR_VAL As Double

Dim APR_VAL As Double
Dim APR_HILO_VAL As Double
Dim APR_HIPC_VAL As Double
Dim APR_LOPC_VAL As Double
Dim APR_FACTOR_VAL As Double

Dim ATR_EMA1_VAL As Double
Dim ATR_EMA2_VAL As Double

Dim APR_EMA1_VAL As Double
Dim APR_EMA2_VAL As Double

Dim COST_VAL As Double
Dim LAST_OPEN_PRICE As Double
Dim LAST_TRADED_PRICE As Double

Dim TICKER_STR As String
Dim TICKERS_VECTOR As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL


'Okay , here 's what we're doing:
'1. Each day we calculate the Percentage Range (PR) of stock prices as the Maximum of:
'* 100*[ (today's High)- (today's Low) ] / (yesterday's Close)
'* 100*| (today's High) / (yesterday's Close) - 1 |
'* 100*| (today's Low) / (yesterday's Close) - 1|
   
'2. Then we calculate the 20-day Exponential Moving Average (EMA) of these daily PRs to
'get the APR, according to:
'APR(today) = (1 - 1/20) APR(yesterday) + (1/20) PR(today)

'3. Then we take 1% of the money we have to invest (our "Equity") and divide by today's
'(APR)*(Stock Price) to get:
'gST = (1% of Equity) / [ (APR)*(Stock Price) ]

'And that gives the number of shares we buy?
'Well, that or a multiple of that. Our trades are 1 gST or 2 gST ... or whatever.
'However, if'n you buy lots of stocks, you'll need lots of money. Here are some DOW stocks
'Note that we calculate APR as a percentage, not a decimal.
'That is, for a 2.5% average/maximum 20-day variation, we take APR = 2.5, not 0.025.

'That COST is what it'd cost to buy that many shares?
'Yes, to buy 1 gST of each, assuming $100K Equity ... so 1% of Equity is $1K.

'So what's the TOTAL cost? you don 't wanna know. ($458K) --> ALL STOCK IN THE DOW :)
'So what does gST stand for? good Stock Trade.

'Some observations:

'* For two assets X and Y with the same price, if the daily volatility of X (measured by APR)
'is twice that of Y, then you'd buy half as many shares.
    
'* If the price of X were twice that of Y, you'd invest the same dollar amount in each.
    
'* If you were considering investing in a basket of stocks (perhaps a few of the DOW stocks),
'then gST would provide an allocation of your Equity.
    
'* There is nothing sacred about a 20-day average. A common average for ATR (or APR?) is 14 days.

'* There is nothing sacred about 1% of your Equity. The gST will provide guidance concerning the
'relative amounts of each asset you'd buy.

'* Note that APR*(Stock Price) gives some estimation of the maximum daily variation in stock price.
'o If APR = 1.2% and Stock Price = $20, then APR*(Stock Price) = 1.2*20 = 24.
'o That suggests that each share of the stock could vary by as much as $0.24 on a given day.

'o If you had gST shares of the stock, that'd imply a daily variation as large as gST*APR*(Stock Price).
'o If you wanted to ensure that such a daily variation didn't exceed 1% of your Equity, then you'd
'want to consider having:
'gST*APR*(Stock Price) = (1% of Equity).
'o That 'd make:
'gST = (1% of Equity) / [ APR*(Stock Price) ]   ... which, of course, is the way we defined gST

'* If you think that this is the way to preserve your financial health, you might also consider tylenol.

'References:
'http://www.financialwebring.org/gummystuff/turtle-trading.htm
'http://www.financialwebring.org/gummystuff/turtle-trading-2.htm
'http://www.financialwebring.org/gummystuff/Donchian.htm

If IsArray(TICKERS_RNG) = True Then
    TICKERS_VECTOR = TICKERS_RNG
    If UBound(TICKERS_VECTOR, 1) = 1 Then
        TICKERS_VECTOR = MATRIX_TRANSPOSE_FUNC(TICKERS_VECTOR)
    End If
Else
    ReDim TICKERS_VECTOR(1 To 1, 1 To 1)
    TICKERS_VECTOR(1, 1) = TICKERS_RNG
End If
NCOLUMNS = UBound(TICKERS_VECTOR, 1)

ReDim TEMP_MATRIX(0 To NCOLUMNS, 1 To 7)
TEMP_MATRIX(0, 1) = "SYMBOLS"

TEMP_MATRIX(0, 2) = EMA_PERIOD & " - PERIOD ATR"
TEMP_MATRIX(0, 3) = "ATR GST/UNITS" 'Good Stock Trade
TEMP_MATRIX(0, 4) = "ATR COST/INVESTMENT"

TEMP_MATRIX(0, 5) = EMA_PERIOD & " - PERIOD APR"
TEMP_MATRIX(0, 6) = "APR GST/UNITS"
TEMP_MATRIX(0, 7) = "APR COST/INVESTMENT"

'COST_VAL = INITIAL_CASH * TURTLE_RULE
'That COST is what it'd cost to buy that many shares?
'Yes, to buy 1 gST of each, assuming $100K Equity ... so 1% of Equity is $1K.
'ATR_FACTOR_VAL = 1
'APR_FACTOR_VAL = 100

COST_VAL = INITIAL_CASH
'ATR --> Dollar Value
ATR_FACTOR_VAL = COST_VAL * TURTLE_RULE / 1
'APR --> Percent
APR_FACTOR_VAL = COST_VAL * TURTLE_RULE / 100

ALPHA_VAL = 1 - 1 / (EMA_PERIOD)

For j = 1 To NCOLUMNS

    TICKER_STR = TICKERS_VECTOR(j, 1)
    TEMP_MATRIX(j, 1) = TICKER_STR
    DATA_MATRIX = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, START_DATE, END_DATE, _
                  "d", "DOHLCVA", False, False, True)
    If IsArray(DATA_MATRIX) = False Then: GoTo 1983
    NROWS = UBound(DATA_MATRIX, 1)
    
    LAST_OPEN_PRICE = DATA_MATRIX(NROWS, 2)
    LAST_TRADED_PRICE = DATA_MATRIX(NROWS, 5)

    i = 2
    GoSub APR_ATR_LINE
    ATR_EMA1_VAL = ALPHA_VAL * 0 + (1 - ALPHA_VAL) * ATR_VAL
    APR_EMA1_VAL = ALPHA_VAL * 0 + (1 - ALPHA_VAL) * APR_VAL
    For i = 3 To NROWS
        GoSub APR_ATR_LINE
        ATR_EMA2_VAL = ALPHA_VAL * ATR_EMA1_VAL + (1 - ALPHA_VAL) * ATR_VAL
        APR_EMA2_VAL = ALPHA_VAL * APR_EMA1_VAL + (1 - ALPHA_VAL) * APR_VAL
        
        ATR_EMA1_VAL = ATR_EMA2_VAL
        APR_EMA1_VAL = APR_EMA2_VAL
    Next i
    
    TEMP_MATRIX(j, 2) = ATR_EMA2_VAL
    TEMP_MATRIX(j, 3) = Round(ATR_FACTOR_VAL / ATR_EMA2_VAL / LAST_TRADED_PRICE, 0)
    TEMP_MATRIX(j, 4) = TEMP_MATRIX(j, 3) * LAST_OPEN_PRICE
    'TEMP_MATRIX(j, 3) = COST_VAL / LAST_TRADED_PRICE / TEMP_MATRIX(j, 2)
'    TEMP_MATRIX(j, 4) = TEMP_MATRIX(j, 3) * LAST_TRADED_PRICE
    
    TEMP_MATRIX(j, 5) = APR_EMA2_VAL
    TEMP_MATRIX(j, 6) = Round(APR_FACTOR_VAL / APR_EMA2_VAL / LAST_TRADED_PRICE, 0)
    TEMP_MATRIX(j, 7) = TEMP_MATRIX(j, 6) * LAST_OPEN_PRICE
    'TEMP_MATRIX(j, 6) = COST_VAL / LAST_TRADED_PRICE / TEMP_MATRIX(j, 5)
'    TEMP_MATRIX(j, 7) = TEMP_MATRIX(j, 6) * LAST_TRADED_PRICE
1983:
Next j

ASSETS_TURTLE_TRADING_GST_FUNC = TEMP_MATRIX

Exit Function
'-----------------------------------------------------------------------------------------------------
APR_ATR_LINE:
'-----------------------------------------------------------------------------------------------------
    
    ATR_HILO_VAL = (DATA_MATRIX(i, 3) - DATA_MATRIX(i, 4)) '/ ATR_FACTOR_VAL
    'ATR_HILO_VAL = ATR_HILO_VAL * ATR_FACTOR_VAL
    ATR_VAL = ATR_HILO_VAL
    
    ATR_HIPC_VAL = Abs(DATA_MATRIX(i, 3) - DATA_MATRIX(i - 1, 5)) '/ ATR_FACTOR_VAL
    'ATR_HIPC_VAL = ATR_HIPC_VAL * ATR_FACTOR_VAL
    If ATR_HIPC_VAL > ATR_VAL Then: ATR_VAL = ATR_HIPC_VAL
    
    ATR_LOPC_VAL = Abs(DATA_MATRIX(i, 4) - DATA_MATRIX(i - 1, 5)) '/ ATR_FACTOR_VAL
    'ATR_LOPC_VAL = ATR_LOPC_VAL * ATR_FACTOR_VAL
    If ATR_LOPC_VAL > ATR_VAL Then: ATR_VAL = ATR_LOPC_VAL

    
    APR_HILO_VAL = (DATA_MATRIX(i, 3) - DATA_MATRIX(i, 4)) / DATA_MATRIX(i - 1, 5) '/ APR_FACTOR_VAL
    'APR_HILO_VAL = APR_HILO_VAL * APR_FACTOR_VAL
    APR_VAL = APR_HILO_VAL
    
    APR_HIPC_VAL = Abs(DATA_MATRIX(i, 3) / DATA_MATRIX(i - 1, 5) - 1) '/ APR_FACTOR_VAL
    'APR_HIPC_VAL = APR_HIPC_VAL * APR_FACTOR_VAL
    If APR_HIPC_VAL > APR_VAL Then: APR_VAL = APR_HIPC_VAL
    
    APR_LOPC_VAL = Abs(DATA_MATRIX(i, 4) / DATA_MATRIX(i - 1, 5) - 1) '/ APR_FACTOR_VAL
    'APR_LOPC_VAL = APR_LOPC_VAL * APR_FACTOR_VAL
    If APR_LOPC_VAL > APR_VAL Then: APR_VAL = APR_LOPC_VAL
'-----------------------------------------------------------------------------------------------------
Return
'-----------------------------------------------------------------------------------------------------
ERROR_LABEL:
'-----------------------------------------------------------------------------------------------------
ASSETS_TURTLE_TRADING_GST_FUNC = Err.number
End Function


Function ASSETS_TURTLE_LONG_SHORT_SIGNAL_FUNC(ByVal TICKERS_RNG As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal END_DATE As Date, _
Optional ByVal ENTRY_BREAKOUT_PERIOD As Long = 20, _
Optional ByVal EMA_SHORT_PERIOD As Long = 8, _
Optional ByVal EMA_LONG_PERIOD As Long = 78)

Dim i As Long
Dim j As Long
Dim k As Long 'BREAKOUT_START_DAY

Dim NROWS As Long
Dim NSIZE As Long

Dim A0_VAL As Double
Dim A1_VAL As Double

Dim B0_VAL As Double
Dim B1_VAL As Double

Dim LOW_VAL As Double 'BREAKOUT_LOW_PRICE
Dim HIGH_VAL As Double 'BREAKOUT_START_DAY

Dim ALPHA1_VAL As Double
Dim ALPHA2_VAL As Double

Dim LONG_FLAG As Boolean
Dim SHORT_FLAG As Boolean

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

ALPHA1_VAL = 1 - 2 / (EMA_SHORT_PERIOD + 1)
ALPHA2_VAL = 1 - 2 / (EMA_LONG_PERIOD + 1)

ReDim TEMP_MATRIX(0 To NSIZE, 1 To 3)
TEMP_MATRIX(0, 1) = "TICKER"
TEMP_MATRIX(0, 2) = "LONG SIGNAL: " & ENTRY_BREAKOUT_PERIOD & "-day Breakout"
TEMP_MATRIX(0, 3) = "SHORT SIGNAL: " & ENTRY_BREAKOUT_PERIOD & "-day Breakout"

For j = 1 To NSIZE
    TICKER_STR = TICKERS_VECTOR(j, 1)
    TEMP_MATRIX(j, 1) = TICKER_STR
    DATA_MATRIX = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, START_DATE, END_DATE, _
                  "d", "DOHLCVA", False, True, True)
    If IsArray(DATA_MATRIX) = False Then: GoTo 1983
    NROWS = UBound(DATA_MATRIX, 1)
    k = NROWS - ENTRY_BREAKOUT_PERIOD
    HIGH_VAL = -2 ^ 52 'initialize the breakout high price
    LOW_VAL = 2 ^ 52 'initialize the breakout high price
    
    i = 1
    A0_VAL = 0: A1_VAL = 0
    B0_VAL = 0: B1_VAL = 0
    
    A0_VAL = (ALPHA1_VAL * DATA_MATRIX(i, 5) + (1 - ALPHA1_VAL) * DATA_MATRIX(i, 5))
    B0_VAL = (ALPHA2_VAL * DATA_MATRIX(i, 5) + (1 - ALPHA2_VAL) * DATA_MATRIX(i, 5))
    
    For i = 2 To NROWS
        A1_VAL = ALPHA1_VAL * A0_VAL + (1 - ALPHA1_VAL) * DATA_MATRIX(i, 5)
        B1_VAL = ALPHA2_VAL * B0_VAL + (1 - ALPHA2_VAL) * DATA_MATRIX(i, 5)
        If i >= k Then
            If DATA_MATRIX(i - 1, 3) > HIGH_VAL Then
                HIGH_VAL = DATA_MATRIX(i - 1, 3)
            ElseIf DATA_MATRIX(i - 1, 3) < LOW_VAL Then
                LOW_VAL = DATA_MATRIX(i - 1, 4)
            End If
            LONG_FLAG = False
            If DATA_MATRIX(i - 1, 5) < HIGH_VAL And DATA_MATRIX(i, 5) > HIGH_VAL Then: LONG_FLAG = True
            'If A1_VAL > B1_VAL And DATA_MATRIX(i - 1, 5) < HIGH_VAL And _
             DATA_MATRIX(i, 5) > HIGH_VAL Then: LONG_FLAG = True 'with LT filter
            SHORT_FLAG = False
            If DATA_MATRIX(i - 1, 5) < LOW_VAL And DATA_MATRIX(i, 5) < LOW_VAL Then: SHORT_FLAG = True
            'If A1_VAL < B1_VAL And DATA_MATRIX(i - 1, 5) < LOW_VAL And _
             DATA_MATRIX(i, 5) < LOW_VAL Then: SHORT_FLAG = True 'with LT filter
        End If
        
        A0_VAL = A1_VAL: B0_VAL = B1_VAL
    Next i
    TEMP_MATRIX(j, 2) = LONG_FLAG
    TEMP_MATRIX(j, 3) = SHORT_FLAG
1983:
Next j

ASSETS_TURTLE_LONG_SHORT_SIGNAL_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
ASSETS_TURTLE_LONG_SHORT_SIGNAL_FUNC = Err.number
End Function


'Turtle Trading System

'----------------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------------

'Each day we calculate TR: the True Range. That 's the maximum value of:
'o (today's High) - (today's Low)
'o | (today's High) - (yesterday's Close) |
'o | (today's Low) - (yesterday's Close) |
    
'Then we calculate the 20-day Exponential Moving Average of the TR.
'We use the following prescription, calling the result N:
'(If it were an ordinary, garden variety moving average rather than an exponential
'moving average, it might also be called Average True Range.
'o N(today) =(19/20) N(yesterday) + (1/20) TR(today)
'Note that we need 20 days worth of data in order to begin our calculations.

'Huh? 19/20 and 1/20? Isn't that the 19-day EMA?
'Yes, but I'm just regurgitating the explanation given in the PDF file ... so don't
'worry about it.

'Why an exponential average ... and why some True Range thing?
'The True Range is a measure of the daily volatility, the price swings over a 24 hour period,
'the degree of violent behaviour - and we want some moving average smooothing, hence EMA so ...

'Okay, armed with the value of N, we calculate a Dollar Volatility like so:
'* Suppose that a 1 point change in the asset generates a change of $D in the contract. (D is
'Dollars per Point.)
'o Dollar Volatility = N D

'Huh? Dollars per Point?
'Yeah, that confused me, too.
'In the Turtle System, practised by Richard Dennis (and his students), they were trading in futures.
'For example, one futures contract for heating oil represents 42,000 gallons. (That's 1000 barrels.)
'Then, for heating oil, a $1 change in the price of heating oil would change the price of the contract
'by $D = $42,000.

'--------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------
'Okay, now suppose you have $1M to invest. How much should you invest in the asset with a given
'Dollar Volatility?
'That was a rhetorical question.
'--------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------
'The Turtle answer is 1% of your equity per unit of Dollar Volatility.
'That is, you build your position in the asset in "Units".
'Each unit is 1% of your equity for each unit of Dollar Volatility.
'In other words, each Unit is:
'Unit = (1% of Portfolio Equity) / (Dollar Volatility)
'--------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------
'Okay, altogether now:
'If N = 20-day Exponential Moving Average of the True Range
'and
'$D is the change in the asset for a $1 change in the underlying commodity
'and
'Dollar Volatility = N D
'then
'Unit = (1% of Portfolio Equity) / (Dollar Volatility)
'--------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------
'Provide an example? Okay, here's an example:
'[1]
'* Suppose you have $1,000,000 to invest in heating oil futures.
'* Assume the Average True Range of heating oil contracts (that's the 20-day EMA of the True Range)
'  is N = 0.015.

'* A $1 point change in the price of oil would generate a change in value of the contract
'  of $D = $42,000.

'* The Dollar Volatility for heating oil contracts is then: N D = (0.015)(42,000) = 630.

'* That makes the size of each trading unit: Unit = (0.01)(1,000,000)/630 = 15.9.

'* Rounding, you might buy 16 contracts.
'--------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------

'$1M to invest? You kidding? Where would I get that kind of money. I'd be lucky if I had ...
'Well, nobody is forcing you to invest in heating oil. You can invest in GE stock with help from
'the Turtle. Remember that the value of the Unit tells you what fraction of 1% of your equity you
'should invest in the stock.

'--------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------
'[2]
'* Suppose you have $10K to invest in GE stock.
'* 1% of that equity is then $100 ... and we want to know what fraction of that should be
'invested in GE stock.

'It 'd be multiples of $100/(Dollar Volatility).

'* Assume the Average True Range of GE stock prices (that's the 20-day EMA of the TR)
'is N = 1.45.

'* For stock (as opposed to futures contracts), we take $D = the price per share of the stock.
'  For GE, that'd be $D = $18.86.

'* Then Dollar Volatility = N D = (1.45)(18.86) = $27.35 per share.
    
'* That makes the size of each trading unit: Unit = (1% of $10K) / (Dollar Volatility) =
'100/27.35 = 3.7 shares.

'Note that this is a ratio of (Dollars)/(Dollars per Share) so the dimensions are
'in Shares... MAYBE!!

'* Rounding, each unit would be 4 shares of GE stock.

'You can trade 2 units or 3, but (according to the Turtle Rules) never more than 4. Of course,
'it's assumed that you have several investments, not just GE.

'-------------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------------

'Recap!

'* We started with the True Range of prices per share (or per futures contract). That's
'measured in Dollars per Share.

'* We averaged these Dollars per Share over the last 20 days. We got N, which is also measured
'in Dollars per Share.

'If N = 1.25, it means the maximimum daily variation in prices, averaged over 20 days,
'is $1.25 per share.

'* Then we introduce $D, the so-called "dollars per point". For our stocks, that'd be measured
'in Dollars per Share.

'* Hence the Dollar Volatility = N D is measured in (gulp!) (Dollars / Share)2.
'* Hence our Unit = (1% of Dollars) / (Dollar Volatility) is measured in (Dollars)/(Dollars / Share)2.

'It seems to me that, at least for trading stock shares, the denominator in the Unit ratio should
'be measured in (Dollars / Share).

'Then the Unit ratio would be measured in Shares.
'To do this, we'd want N to be a percentage ... like, maybe, (20-day average True Range of prices
'per share) divided by the (Price per Share).

'Then the Unit ratio would be (Dollars) / (Dollars / Share) ... or Shares.

'In examples I've found, they're talking about commodities.

'For example [1], a heating oil contract in March, 2003 was trading at about $0.60 dollars per
'contract and the average True Range was about 0.015 dollars per contract, as we noted above.

'(That's the example in the PDF paper I mentioned earlier. It's saying that the 20-day average of
'"price variations" was $0.015 per contract)

'Then the Unit ratio would be (Dollars) / (Dollars / Contract)2 and that ain't right.
'Indeed, that PDF paper says that the Unit ratio is in "contracts".


'Following the "typical" Turtle System, we take the m-day EMA using the magic formula:
'N(today) =(1 - 1/m) N(yesterday)+ (1/m)TR(today) .
'That 's give weights 1 - 1/m = (19/20) and 1/m = (1/20) for m = 20.
'Until further notice (!), the Unit ratio is based upon $100K equity (so 1% = $1K).
'Hence Unit = $1000/[N*(Current Price)].
'In the spreadsheet example, that'd be: 1000/[1.47*11.82] = 57.55.

'>Is that 57.55 shares ... or what?
'i 'm still cogitating ...
'I think it's reasonable to modify the above spreadsheet so that I calculate the 20-day APR
'rather than the 20-day ATR.

'Don 't you remember? We talked about that here.
'That 's the 20-day Average Percentage Range where the Percentage Range is defined like so:
'* Each day you calculate the largest of the following numbers:
'1. [ (today's High)- (today's Low) ] / (yesterday's Close)
'2. | (today's High) / (yesterday's Close) - 1 |
'3. | (today's Low) / (yesterday's Close) - 1|
'* Called these numbers Percentage Ranges (or PR).
'* You then average these "Percentage Ranges" over the past m days, calling it the Average
'Percentage Range (or APR).

'Since the APR is a percentage (it's some price range divided by yesterday's price), then our Unit
'will be measured in Shares.
'Then we're happy.


