Attribute VB_Name = "FINAN_PORT_TRADE_LOGS_LIBR"

'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------

'************************************************************************************
'************************************************************************************
'FUNCTION      : PORTFOLIO_TRANSACTIONS_LOG_FUNC
'DESCRIPTION   :
'LIBRARY       : PORT_TRADE
'GROUP         : TRACK
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 22/07/2010
'************************************************************************************
'************************************************************************************

Function PORTFOLIO_TRANSACTIONS_LOG_FUNC( _
ByRef TX_ID_RNG As Variant, _
ByRef DATE_RNG As Variant, _
ByRef ACTION_RNG As Variant, _
ByRef SHARES_RNG As Variant, _
ByRef SECURITY_RNG As Variant, _
ByRef UNDERLYING_RNG As Variant, _
ByRef TYPE_RNG As Variant, _
ByRef CASH_FLOW_RNG As Variant, _
ByRef PRICE_RNG As Variant, _
ByRef COMMISION_RNG As Variant, _
ByRef ACCOUNT_RNG As Variant, _
ByRef INV_AMT_RNG As Variant, _
ByRef CASH_AMT_RNG As Variant, _
ByRef CASH_ACCT_RNG As Variant, _
ByRef NOTES_RNG As Variant, _
Optional ByVal OUTPUT As Integer = 1)

'TxID: TX_ID_RNG
'Date: DATE_RNG
'Action: ACTION_RNG
'B --> Buy | BTC --> Buy to Cover | Div --> Dividend | DivX --> Reinvested Dividend |
'Int --> Interest Income | S --> Sell | SS --> Sold Short | CA --> Call Assigned |
'CE --> Call Expired | C --> Called | CC --> Call Closed | PA --> Put Assigned |
'PE --> Put Expired | PC --> Put Closed | T --> Transfer | A --> Adjustment |
'R --> Reorganization

'Shares: SHARES_RNG

'SECURITY: SECURITY_RNG
'Cash,Stock,Mutual Fund,Taxable Bond Fund,Tax Free Bond Fund,Option,Real Estate

'Underlying: UNDERLYING_RNG

'Type: TYPE_RNG
'SECURITY Type: Cash,Stock,Option,Bond...

'Cash Flow: CASH_FLOW_RNG
'Price: PRICE_RNG
'Commision: COMMISION_RNG

'Account: ACCOUNT_RNG
'401K,UGMA,IRA,Personal

'Inv Amt: INV_AMT_RNG
'Cash Amt: CASH_AMT_RNG
'Cash Acct: CASH_ACCT_RNG
'Notes: NOTES_RNG

Dim i As Long
Dim j As Long
Dim k As Long
Dim NROWS As Long

Dim KEY_STR As String
Dim HEADINGS_ARR As Variant
Dim TEMP_OBJ As New Collection

Dim TX_ID_VECTOR As Variant
Dim DATE_VECTOR As Variant
Dim ACTION_VECTOR As Variant
Dim SHARES_VECTOR As Variant

Dim SECURITY_VECTOR As Variant
Dim UNDERLYING_VECTOR As Variant
Dim TYPE_VECTOR As Variant

Dim CASH_FLOW_VECTOR As Variant
Dim PRICE_VECTOR As Variant
Dim COMMISION_VECTOR As Variant
Dim ACCOUNT_VECTOR As Variant
Dim INV_AMT_VECTOR As Variant
Dim CASH_AMT_VECTOR As Variant
Dim CASH_ACCT_VECTOR As Variant
Dim NOTES_VECTOR As Variant

Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

TX_ID_VECTOR = TX_ID_RNG
If UBound(TX_ID_VECTOR) = 1 Then
    TX_ID_VECTOR = MATRIX_TRANSPOSE_FUNC(TX_ID_VECTOR)
End If
NROWS = UBound(TX_ID_VECTOR, 1)

DATE_VECTOR = DATE_RNG
If UBound(DATE_VECTOR) = 1 Then
    DATE_VECTOR = MATRIX_TRANSPOSE_FUNC(DATE_VECTOR)
End If
If UBound(DATE_VECTOR, 1) <> NROWS Then: GoTo ERROR_LABEL

ACTION_VECTOR = ACTION_RNG
If UBound(ACTION_VECTOR) = 1 Then
    ACTION_VECTOR = MATRIX_TRANSPOSE_FUNC(ACTION_VECTOR)
End If
If UBound(ACTION_VECTOR, 1) <> NROWS Then: GoTo ERROR_LABEL

SHARES_VECTOR = SHARES_RNG
If UBound(SHARES_VECTOR) = 1 Then
    SHARES_VECTOR = MATRIX_TRANSPOSE_FUNC(SHARES_VECTOR)
End If
If UBound(SHARES_VECTOR, 1) <> NROWS Then: GoTo ERROR_LABEL

SECURITY_VECTOR = SECURITY_RNG
If UBound(SECURITY_VECTOR) = 1 Then
    SECURITY_VECTOR = MATRIX_TRANSPOSE_FUNC(SECURITY_VECTOR)
End If
If UBound(SECURITY_VECTOR, 1) <> NROWS Then: GoTo ERROR_LABEL

UNDERLYING_VECTOR = UNDERLYING_RNG
If UBound(UNDERLYING_VECTOR) = 1 Then
    UNDERLYING_VECTOR = MATRIX_TRANSPOSE_FUNC(UNDERLYING_VECTOR)
End If
If UBound(UNDERLYING_VECTOR, 1) <> NROWS Then: GoTo ERROR_LABEL

TYPE_VECTOR = TYPE_RNG
If UBound(TYPE_VECTOR) = 1 Then
    TYPE_VECTOR = MATRIX_TRANSPOSE_FUNC(TYPE_VECTOR)
End If
If UBound(TYPE_VECTOR, 1) <> NROWS Then: GoTo ERROR_LABEL

CASH_FLOW_VECTOR = CASH_FLOW_RNG
If UBound(CASH_FLOW_VECTOR) = 1 Then
    CASH_FLOW_VECTOR = MATRIX_TRANSPOSE_FUNC(CASH_FLOW_VECTOR)
End If
If UBound(CASH_FLOW_VECTOR, 1) <> NROWS Then: GoTo ERROR_LABEL

PRICE_VECTOR = PRICE_RNG
If UBound(PRICE_VECTOR) = 1 Then
    PRICE_VECTOR = MATRIX_TRANSPOSE_FUNC(PRICE_VECTOR)
End If
If UBound(PRICE_VECTOR, 1) <> NROWS Then: GoTo ERROR_LABEL

COMMISION_VECTOR = COMMISION_RNG
If UBound(COMMISION_VECTOR) = 1 Then
    COMMISION_VECTOR = MATRIX_TRANSPOSE_FUNC(COMMISION_VECTOR)
End If
If UBound(COMMISION_VECTOR, 1) <> NROWS Then: GoTo ERROR_LABEL

ACCOUNT_VECTOR = ACCOUNT_RNG
If UBound(ACCOUNT_VECTOR) = 1 Then
    ACCOUNT_VECTOR = MATRIX_TRANSPOSE_FUNC(ACCOUNT_VECTOR)
End If
If UBound(ACCOUNT_VECTOR, 1) <> NROWS Then: GoTo ERROR_LABEL

INV_AMT_VECTOR = INV_AMT_RNG
If UBound(INV_AMT_VECTOR) = 1 Then
    INV_AMT_VECTOR = MATRIX_TRANSPOSE_FUNC(INV_AMT_VECTOR)
End If
If UBound(INV_AMT_VECTOR, 1) <> NROWS Then: GoTo ERROR_LABEL

CASH_AMT_VECTOR = CASH_AMT_RNG
If UBound(CASH_AMT_VECTOR) = 1 Then
    CASH_AMT_VECTOR = MATRIX_TRANSPOSE_FUNC(CASH_AMT_VECTOR)
End If
If UBound(CASH_AMT_VECTOR, 1) <> NROWS Then: GoTo ERROR_LABEL

CASH_ACCT_VECTOR = CASH_ACCT_RNG
If UBound(CASH_ACCT_VECTOR) = 1 Then
    CASH_ACCT_VECTOR = MATRIX_TRANSPOSE_FUNC(CASH_ACCT_VECTOR)
End If
If UBound(CASH_ACCT_VECTOR, 1) <> NROWS Then: GoTo ERROR_LABEL

NOTES_VECTOR = NOTES_RNG
If UBound(NOTES_VECTOR) = 1 Then
    NOTES_VECTOR = MATRIX_TRANSPOSE_FUNC(NOTES_VECTOR)
End If
If UBound(NOTES_VECTOR, 1) <> NROWS Then: GoTo ERROR_LABEL

'--------------------------------------------------------------------------------------
Select Case OUTPUT
'--------------------------------------------------------------------------------------
Case 0
'--------------------------------------------------------------------------------------
    HEADINGS_ARR = Array("TXID", "DATE", "ACTION", "SHARES", "SECURITY", _
                         "UNDERLYING", "CASH FLOW", "PRICE", "COMMISION", _
                         "ACCOUNT", "INV AMT", "CASH AMT", "CASH ACCT", _
                         "NOTES", "PORTFOLIO/SECURITY", _
                         "PORTFOLIO/UNDERLYING", _
                         "TYPE", "PORTFOLIO/CASH")
    ReDim TEMP_MATRIX(0 To NROWS, 1 To 18)
    i = LBound(HEADINGS_ARR)
    For j = 1 To 18
        TEMP_MATRIX(0, j) = HEADINGS_ARR(i)
        i = i + 1
    Next j
    For i = 1 To NROWS
        TEMP_MATRIX(i, 1) = TX_ID_VECTOR(i, 1)
        TEMP_MATRIX(i, 2) = DATE_VECTOR(i, 1)
        TEMP_MATRIX(i, 3) = ACTION_VECTOR(i, 1)
        TEMP_MATRIX(i, 4) = SHARES_VECTOR(i, 1)
        TEMP_MATRIX(i, 5) = SECURITY_VECTOR(i, 1)
        TEMP_MATRIX(i, 6) = UNDERLYING_VECTOR(i, 1)
        TEMP_MATRIX(i, 7) = CASH_FLOW_VECTOR(i, 1)
        TEMP_MATRIX(i, 8) = PRICE_VECTOR(i, 1)
        TEMP_MATRIX(i, 9) = COMMISION_VECTOR(i, 1)
        TEMP_MATRIX(i, 10) = ACCOUNT_VECTOR(i, 1)
        TEMP_MATRIX(i, 11) = INV_AMT_VECTOR(i, 1)
        TEMP_MATRIX(i, 12) = CASH_AMT_VECTOR(i, 1)
        TEMP_MATRIX(i, 13) = CASH_ACCT_VECTOR(i, 1)
        TEMP_MATRIX(i, 14) = NOTES_VECTOR(i, 1)
        
        TEMP_MATRIX(i, 15) = TEMP_MATRIX(i, 10) & " / " & TEMP_MATRIX(i, 5)
        TEMP_MATRIX(i, 16) = TEMP_MATRIX(i, 10) & " / " & TEMP_MATRIX(i, 6)
        TEMP_MATRIX(i, 17) = TYPE_VECTOR(i, 1)
        TEMP_MATRIX(i, 18) = TEMP_MATRIX(i, 10) & " / " & TEMP_MATRIX(i, 13)
    Next i
    PORTFOLIO_TRANSACTIONS_LOG_FUNC = TEMP_MATRIX
'--------------------------------------------------------------------------------------
Case Else 'Summary per Account
'--------------------------------------------------------------------------------------
    Err.Clear
    On Error Resume Next
    j = 0
    For i = 1 To NROWS
        If ACCOUNT_VECTOR(i, 1) <> "" Then
            KEY_STR = ACCOUNT_VECTOR(i, 1) & " / " & SECURITY_VECTOR(i, 1)
            Call TEMP_OBJ.Add(KEY_STR, KEY_STR)
            If Err.number = 0 Then
                j = j + 1
            Else
                Err.Clear
            End If
        End If
    Next i
    ReDim TEMP_MATRIX(0 To j, 1 To 6)
    TEMP_MATRIX(0, 1) = "Account / Security"
    TEMP_MATRIX(0, 2) = "Shares"
    TEMP_MATRIX(0, 3) = "Cash Flow"
    TEMP_MATRIX(0, 4) = "Cost Per Share"
    TEMP_MATRIX(0, 5) = "Adjusted Cash Flow"
    TEMP_MATRIX(0, 6) = "Adjusted Cost Per Share"
    For k = 1 To j
        TEMP_MATRIX(k, 1) = TEMP_OBJ(k)
        TEMP_MATRIX(k, 2) = 0
        For i = 1 To NROWS
            KEY_STR = ACCOUNT_VECTOR(i, 1) & " / " & SECURITY_VECTOR(i, 1)
            If KEY_STR = TEMP_MATRIX(k, 1) Then
                TEMP_MATRIX(k, 2) = TEMP_MATRIX(k, 2) + SHARES_VECTOR(i, 1)
            End If
            
            KEY_STR = ACCOUNT_VECTOR(i, 1) & " / " & CASH_ACCT_VECTOR(i, 1)
            If KEY_STR = TEMP_MATRIX(k, 1) Then 'Is the account type = Cash????
                TEMP_MATRIX(k, 3) = TEMP_MATRIX(k, 3) - CASH_FLOW_VECTOR(i, 1)
                TEMP_MATRIX(k, 5) = TEMP_MATRIX(k, 5) - CASH_FLOW_VECTOR(i, 1)
            Else
                KEY_STR = ACCOUNT_VECTOR(i, 1) & " / " & SECURITY_VECTOR(i, 1)
                If KEY_STR = TEMP_MATRIX(k, 1) Then
                    TEMP_MATRIX(k, 3) = TEMP_MATRIX(k, 3) + CASH_FLOW_VECTOR(i, 1)
                End If
                KEY_STR = ACCOUNT_VECTOR(i, 1) & " / " & UNDERLYING_VECTOR(i, 1)
                If KEY_STR = TEMP_MATRIX(k, 1) Then
                    TEMP_MATRIX(k, 5) = TEMP_MATRIX(k, 5) + CASH_FLOW_VECTOR(i, 1)
                End If
            End If
        Next i
        If TEMP_MATRIX(k, 2) <> 0 Then
            TEMP_MATRIX(k, 4) = -TEMP_MATRIX(k, 3) / TEMP_MATRIX(k, 2)
            TEMP_MATRIX(k, 6) = -TEMP_MATRIX(k, 5) / TEMP_MATRIX(k, 2)
        Else
            TEMP_MATRIX(k, 4) = "" 'CVErr(xlErrNA)
            TEMP_MATRIX(k, 6) = "" 'CVErr(xlErrNA)
        End If
        
    Next k
    PORTFOLIO_TRANSACTIONS_LOG_FUNC = TEMP_MATRIX
'--------------------------------------------------------------------------------------
End Select
'--------------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
PORTFOLIO_TRANSACTIONS_LOG_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_TRACK_TRADING_ACTIVITIES_FUNC
'DESCRIPTION   :
'This function is designed to help you keep track of your trading activities.
'The function is broken into three main parts, the strategy block, the statistics
'block and the actual trading log.

'Max High: The highest dollar returning trade.

'Max Low: The most negative returning trade.

'# Wins: The number of winning trades.

'# Losses: The number of losing trades.

'Profit Target: This is something for you to enter. Some people like to place
'two additional orders when they enter the market, a 'sell limit' (profit) and
'a 'stop loss' order. This field allows you to do 'what if' scenarios and have
'the 'sell limit' value calculated for you.

'Positions: This value will determine how much of your available trading funds
'will be allocated per position. For example, if this value is one (1), 100% of
'your trading value will be allocated for that one trade. If the value is two (2),
'50% of available funds will be allocated per trade. And so forth.

'Max Risk: You determine how much of your portfolio you want to 'risk' on
'each trade. This field is used to calculate the Suggested Trailing Stop and the Stop.

'Avg.win: Your average win.

'Avg.loss: Your average loss.

'Max Shares: Here we get a bit tricky. Perhaps you are a conservative investor,
'just getting his feet wet, and you have a gut wrenching feeling when you see
'you are about to buy 68,700 shares of Berkshire Hathaway (there's a joke here -
'look up Berkshire Hathaway. If you can buy that much, you sure don't need this
'spreadsheet!). Not to worry, you can use this field to select 100 shares, 200
'shares, you name it. This will limit your total exposure to that many shares.
'Period.

'Current Positions Open: This just keeps track (a double check) of how many
'positions you have open.

'DDLR Ratio: Discipline, Direction, Leverage, Risk Ratio, a term coined by Robert
'Deel in "The Strategic Electronic Day Trader". This value should be above 3 to
'give you a warm and fuzzy feeling about using margin.

'W/L Ratio: How many wins versus how many losses.

'Margin %: If you have a margin account (meaning you can sell short), you can
'specify how much of that available margin you wish to use for trading. I
'recommend not using 100% (or 400% if you are a day trader) as it provides
'you with no room if the position turns against you. Can you say "Margin Call"?

'Perf/Day: This field keeps track of your rate of return on a daily basis,

'Days In System: How many days has this particular trade log existed?

'total Margin: How many of the broker's dollars do you have access to? This is
'an updated value as your account balance increases.

'Performance: The total performance of this trade log for the time it's been
'in effect.

'Cost/Trade: This is the commission. Trade Log treats this value as times two
'since you pay this much going into the trade and this much again when you exit.

'Beg. Bal.: This is your initial portfolio total amount.

'Initial Margin: The amount of margin you are committing on your first trade.

'$ Ret.: The number of dollars this strategy returned from start date to current.

'Account Value: The current value (after all profits and losses are summed) of
'your portfolio.

'Amount Invested: The total investment of all open positions.

'trading Log: This trading log is primarily an EOD trading log; which means you
'will be analyzing signals and/or trading opportunities tonight, determining
'position sizing from available data, and placing the trade, stop loss and
'profit limit tomorrow, in the morning at the open of the market.

'Sym: The stock 's symbol.

'Date in: Tomorrows date (or the next trading days date).

'Prev Close: Required to determine position sizing. It is the close at the end of
'trading today.

'Price In: This value will be filled in AFTER you have made the trade, the trade
'is executed and you have a reported fill price. If you're paper trading, you
'can make this value anything you want.

'Side: There are only two possibilities here - Long or Short. A drop down widow
'will allow you to choose.

'Price Out: This is the price you got when you exited the trade, from whatever
'means. Your exit could have been caused by a stop or limit execution or because
'you felt like getting out of the trade.

'% G/L on Trade: How the trade did. Returned automatically.

'Shares: To recapitulate - you enter the price of the close tonight, and the
'number of shares is calculated for you based on margin, maximum allowed
'shares, account balance, positions open, etc. Note that the price in amount
'may have warranted more or less shares, but the idea is to have your order
'ready to go with but a finger press at the open, or to have placed the order
'after hours before you know what the fill price will be. There rises a problem
'at this point. Suppose you have entered in two trades and the first trade
'reaches a profit limit and exits. You then update the trade log by entering
'in the price out field. This causes your account balance to be recalculated
'and - behold - the number of shares in the second trade has just increased
'as a result of the profit of the first trade, This can make it a bit difficult
'to manage your portfolio properly. Therefore, YOU MUST - let me say that again,
'YOU MUST hardcode the shares value in the shares field. So if the calculated
'amount is 220 shares when you are doing your aftermarket analysis and preparing
'the next days order, you MUST then type in the amount of 220. This will destroy
'the calculation for that field and it will remain constant. This is a good.

'G/L: How the trade did after commissions.

'Profit/Loss: A running total of how your profit (hopefully) is doing.

'Amount Invested: This field is only populated while the trade is actually
'alive. It disappears once the trade is closed.

'Target: This is the amount to use for a limit sell order (if long) based
'upon the Profit Target field in the statistics block.

'Suggested Trailing Stop: A value determined by the Max Risk field in the
'statistics block. Only useful if you use trailing stops.

'Stop: The suggested fixed stop, again based on the Max Risk field in the
'statistics block. Note that this field may be changed as the trade progresses.
'It will affect the Risk field (covered next).

'Risk: This value shows you, in dollars, about how much you may lose on the
'trade if you are using the calculated stops (slippage is not accounted for).
'As the trade progresses and the stop field is manually changed, this field
'will be updated to reflect the new risk (or locked in profit).

'Date Out: A manually entered value of the date the trade was exited.
'Used to calculate the next field.

'Days In Trade: Some people find it useful to know how many days the trade
'was in effect.

'Reference: www.gummy-stuff.org/TradeLog.doc

'LIBRARY       : TRADE
'GROUP         : TRACK
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 03/16/2009
'************************************************************************************
'************************************************************************************

Function PORT_TRACK_TRADING_ACTIVITIES_FUNC(ByRef SYMBOL_RNG As Variant, _
ByRef SETTLEMENT_RNG As Variant, _
ByRef PREV_CLOSE_RNG As Variant, _
ByRef PRICE_IN_RNG As Variant, _
ByRef POSITION_RNG As Variant, _
Optional ByRef SHARES_RNG As Variant, _
Optional ByRef DAY_OUT_RNG As Variant, _
Optional ByRef PRICE_OUT_RNG As Variant, _
Optional ByVal BEG_BALANCE As Double = 10000, _
Optional ByVal PROFIT_TARGET As Double = 0.08, _
Optional ByVal NO_POSITIONS As Long = 10, _
Optional ByVal MAX_RISK As Double = 0.03, _
Optional ByVal MAX_SHARES As Double = 1000, _
Optional ByVal MARGIN_PERCENTAGE As Double = 0.01, _
Optional ByVal COST_TRADE As Double = 8.4, _
Optional ByVal COUNT_BASIS As Double = 252, _
Optional ByVal OUTPUT As Integer = 0, _
Optional ByRef HOLIDAYS_RNG As Variant)

'---------------------------------------------------------------------------------------
'PREV_CLOSE_RNG: Only used for calculating position size for upcoming trade.

'PRICE_IN_RNG: Completing this field (at time of entry) calculates the target value.

'PRICE_OUT_RNG: The price of the stock  when you sold it, or the last price (for
'tracking P/L). You must manually enter this value.
'---------------------------------------------------------------------------------------
'BEG_BALANCE: Your beginning balance

'PROFIT_TARGET: This is the value used to calculate the exit target

'NO_POSITIONS: This value is changed for each trade to reflect how many
'trades there are for the day. Remember, you MUST hardcode the shares bought
'for each position or the number of shares for closed positions will be
'recalculated, giving erroneous results. Change this value to the number
'of positions entered each day.

'MAX_RISK: This is a number you enter - it tells you where to place a
'trailing stop to limit your risk to your tolerance level.

'MAX_SHARES: The maximum number of shares you will have in any position.
'You can change this number to reflect a more conservative approach.
'Type in 100 for the maximum number of shares and the amount of your
'investment will be significantly reduced. NOTE: This value is used only
'for initial calculation. Once the shares field is hardcoded, the shares
'field will not be updated.

'MARGIN_PERCENTAGE: With a 'normal' account, the maximum that margin can
'be is 100%. Keep in mind that as a Pattern Day Trader, this value can be
'as high as 400%. However, with most brokerages, you CAN NOT hold a position
'overnight with more than 100% of margin used or you will get a Reg T Margin
'Call in the morning.

'COST_TRADE: This is the one way cost per trade. Remember, there is this cost
'for entering the trade and the same cost again for exiting the trade. If you
'use IB or Cybertrader, you will need to calculate this differently as a price
'per function of shares purchased.

'COUNT_BASIS: Trading Days per year
'---------------------------------------------------------------------------------------

Dim i As Long
Dim NROWS As Long

Dim TEMP_SUM As Double
Dim TEMP_MAX As Double
Dim TEMP_MIN As Double

Dim DAY_MIN As Date
Dim DAY_MAX As Date

Dim COUNT_LOSS As Long
Dim COUNT_WIN As Long

Dim SUM_POSITIONS As Double
Dim COUNT_POSITIONS As Long

Dim SUM_WIN As Double
Dim SUM_LOSS As Double

Dim TRADING_TABLE As Variant
Dim TRADING_SUMMARY As Variant

Dim SYMBOL_VECTOR As Variant
Dim SETTLEMENT_VECTOR As Variant
Dim PREV_CLOSE_VECTOR As Variant
Dim PRICE_IN_VECTOR As Variant
Dim POSITION_VECTOR As Variant
Dim SHARES_VECTOR As Variant
Dim PRICE_OUT_VECTOR As Variant
Dim DAY_OUT_VECTOR As Variant

Dim lower_date As Date
Dim upper_date As Date

Dim lower_limit As Double
Dim upper_limit As Double

On Error GoTo ERROR_LABEL

lower_date = DateSerial(1950, 1, 1)
upper_date = DateSerial(2050, 1, 1)

lower_limit = -1E+15
upper_limit = 1E+15

SYMBOL_VECTOR = SYMBOL_RNG
    If UBound(SYMBOL_VECTOR, 1) = 1 Then: _
        SYMBOL_VECTOR = MATRIX_TRANSPOSE_FUNC(SYMBOL_VECTOR)

SETTLEMENT_VECTOR = SETTLEMENT_RNG
    If UBound(SETTLEMENT_VECTOR, 1) = 1 Then: _
        SETTLEMENT_VECTOR = MATRIX_TRANSPOSE_FUNC(SETTLEMENT_VECTOR)
If UBound(SYMBOL_VECTOR, 1) <> UBound(SETTLEMENT_VECTOR, 1) Then: GoTo ERROR_LABEL

PREV_CLOSE_VECTOR = PREV_CLOSE_RNG
    If UBound(PREV_CLOSE_VECTOR, 1) = 1 Then: _
        PREV_CLOSE_VECTOR = MATRIX_TRANSPOSE_FUNC(PREV_CLOSE_VECTOR)
If UBound(SYMBOL_VECTOR, 1) <> UBound(PREV_CLOSE_VECTOR, 1) Then: GoTo ERROR_LABEL

PRICE_IN_VECTOR = PRICE_IN_RNG
    If UBound(PRICE_IN_VECTOR, 1) = 1 Then: _
        PRICE_IN_VECTOR = MATRIX_TRANSPOSE_FUNC(PRICE_IN_VECTOR)
If UBound(SYMBOL_VECTOR, 1) <> UBound(PRICE_IN_VECTOR, 1) Then: GoTo ERROR_LABEL
        
POSITION_VECTOR = POSITION_RNG
    If UBound(POSITION_VECTOR, 1) = 1 Then: _
        POSITION_VECTOR = MATRIX_TRANSPOSE_FUNC(POSITION_VECTOR)
If UBound(SYMBOL_VECTOR, 1) <> UBound(POSITION_VECTOR, 1) Then: GoTo ERROR_LABEL

If IsArray(SHARES_RNG) = True Then
    SHARES_VECTOR = SHARES_RNG
    If UBound(SHARES_VECTOR, 1) = 1 Then: _
        SHARES_VECTOR = MATRIX_TRANSPOSE_FUNC(SHARES_VECTOR)
    If UBound(SYMBOL_VECTOR, 1) <> UBound(SHARES_VECTOR, 1) Then: GoTo ERROR_LABEL
Else
    ReDim SHARES_VECTOR(1 To UBound(SYMBOL_VECTOR, 1), 1 To 1)
    For i = 1 To UBound(SHARES_VECTOR, 1)
        SHARES_VECTOR(i, 1) = ""
    Next i
End If

If IsArray(PRICE_OUT_RNG) = True Then
    PRICE_OUT_VECTOR = PRICE_OUT_RNG
    If UBound(PRICE_OUT_VECTOR, 1) = 1 Then: _
        PRICE_OUT_VECTOR = MATRIX_TRANSPOSE_FUNC(PRICE_OUT_VECTOR)
    If UBound(SYMBOL_VECTOR, 1) <> UBound(PRICE_OUT_VECTOR, 1) Then: GoTo ERROR_LABEL
Else
    ReDim PRICE_OUT_VECTOR(1 To UBound(SYMBOL_VECTOR, 1), 1 To 1)
    For i = 1 To UBound(SYMBOL_VECTOR, 1)
        PRICE_OUT_VECTOR(i, 1) = ""
    Next i
End If


If IsArray(DAY_OUT_RNG) = True Then
    DAY_OUT_VECTOR = DAY_OUT_RNG
    If UBound(DAY_OUT_VECTOR, 1) = 1 Then: _
        DAY_OUT_VECTOR = MATRIX_TRANSPOSE_FUNC(DAY_OUT_VECTOR)
    If UBound(SYMBOL_VECTOR, 1) <> UBound(DAY_OUT_VECTOR, 1) Then: GoTo ERROR_LABEL
Else
    ReDim DAY_OUT_VECTOR(1 To UBound(SYMBOL_VECTOR, 1), 1 To 1)
    For i = 1 To UBound(SYMBOL_VECTOR, 1)
        DAY_OUT_VECTOR(i, 1) = ""
    Next i
End If

NROWS = UBound(SYMBOL_VECTOR, 1)
ReDim TRADING_TABLE(0 To NROWS, 1 To 17)

TRADING_TABLE(0, 1) = "SYMBOL"
TRADING_TABLE(0, 2) = "DATE IN"
TRADING_TABLE(0, 3) = "PREVIOUS CLOSE"
TRADING_TABLE(0, 4) = "PRICE IN"
TRADING_TABLE(0, 5) = "POSITION"
TRADING_TABLE(0, 6) = "PRICE OUT"

TRADING_TABLE(0, 7) = "%G/L TRADE" 'The percentage gain or loss for the trade. Note
'that the cost of commissions is NOT figured in this result This is the %G/L for
'the trade..

TRADING_TABLE(0, 8) = "SHARES" 'Once this field is calculated, it MUST be hardcoded
'to the calculated value. Otherwise changes in previous, open trades will cause
'this value to be recalculated.

TRADING_TABLE(0, 9) = "G/L TRADE" 'The profit or loss of the trade. Note that
'commission costs are included in this result.

TRADING_TABLE(0, 10) = "PROFIT LOSS" 'This is a running total of the current
'result of your trading.

TRADING_TABLE(0, 11) = "AMOUNT INVESTED" 'This field is populated only
'during the time the trade (or any portion of) is active.

TRADING_TABLE(0, 12) = "TARGET" 'Target is determined by the value in
'"Profit Target" above.

TRADING_TABLE(0, 13) = "TRAILING STOP" 'The 'suggested' trailing stop, as a
'function of your value for risk.

TRADING_TABLE(0, 14) = "STOP" 'What you should set a fixed stop at as a
'function of your value for risk.

TRADING_TABLE(0, 15) = "RISK" 'This is how much you are risking, not
'including any slippage. Includes trade cost. Populated at "Price In" time.

TRADING_TABLE(0, 16) = "DATE OUT" 'Day the trade is closed.

TRADING_TABLE(0, 17) = "DAYS IN TRADE" 'Days the trade lived. Useful if
'you want to track your trade length.

TEMP_SUM = 0  'INIT_PROFIT
SUM_WIN = 0
SUM_LOSS = 0

COUNT_LOSS = 0
COUNT_WIN = 0

SUM_POSITIONS = 0
COUNT_POSITIONS = 0

DAY_MAX = lower_date
DAY_MIN = upper_date

TEMP_MAX = lower_limit
TEMP_MIN = upper_limit

For i = 1 To NROWS

    TRADING_TABLE(i, 1) = _
        IIf(SYMBOL_VECTOR(i, 1) = "", "", SYMBOL_VECTOR(i, 1))
    
    TRADING_TABLE(i, 2) = _
        IIf(SETTLEMENT_VECTOR(i, 1) = "", "", SETTLEMENT_VECTOR(i, 1))

    If TRADING_TABLE(i, 2) <> "" Then
        If TRADING_TABLE(i, 2) > DAY_MAX Then: DAY_MAX = TRADING_TABLE(i, 2)
        If TRADING_TABLE(i, 2) < DAY_MIN Then: DAY_MIN = TRADING_TABLE(i, 2)
    End If

    TRADING_TABLE(i, 3) = _
        IIf(PREV_CLOSE_VECTOR(i, 1) = "", "", PREV_CLOSE_VECTOR(i, 1))
    
    TRADING_TABLE(i, 4) = _
        IIf(PRICE_IN_VECTOR(i, 1) = "", "", PRICE_IN_VECTOR(i, 1))
    
    TRADING_TABLE(i, 5) = _
        IIf(POSITION_VECTOR(i, 1) = "", "", POSITION_VECTOR(i, 1))
    
    TRADING_TABLE(i, 6) = _
        IIf(PRICE_OUT_VECTOR(i, 1) = "", "", PRICE_OUT_VECTOR(i, 1))
    
    TRADING_TABLE(i, 7) = _
        IIf(PRICE_OUT_VECTOR(i, 1) = "", "", PRICE_OUT_VECTOR(i, 1))
    
    If TRADING_TABLE(i, 6) <> "" Then
        TRADING_TABLE(i, 16) = _
            IIf(DAY_OUT_VECTOR(i, 1) = "", "", DAY_OUT_VECTOR(i, 1))
    
        If (TRADING_TABLE(i, 5) = 1) Then
            TRADING_TABLE(i, 7) = (TRADING_TABLE(i, 6) - _
                TRADING_TABLE(i, 4)) / TRADING_TABLE(i, 4)
        Else 'Short or -1
            TRADING_TABLE(i, 7) = (TRADING_TABLE(i, 4) - _
                TRADING_TABLE(i, 6)) / TRADING_TABLE(i, 4)
        End If
    
    Else
        TRADING_TABLE(i, 16) = ""
        
        TRADING_TABLE(i, 7) = ""
    End If

'-------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------

    If SHARES_VECTOR(i, 1) <> "" Then
        TRADING_TABLE(i, 8) = SHARES_VECTOR(i, 1)
    Else
        If TRADING_TABLE(i, 3) <> "" Then
            If ROUND_FUNC((BEG_BALANCE + TEMP_SUM + _
                          (BEG_BALANCE * MARGIN_PERCENTAGE) + _
                          (MARGIN_PERCENTAGE * TEMP_SUM)) / TRADING_TABLE(i, 3) _
                          / NO_POSITIONS, -1, 4) < MAX_SHARES Then
        
                TRADING_TABLE(i, 8) = ROUND_FUNC((BEG_BALANCE + _
                          TEMP_SUM + (BEG_BALANCE * MARGIN_PERCENTAGE) + _
                          (MARGIN_PERCENTAGE * TEMP_SUM)) / TRADING_TABLE(i, 3) _
                          / NO_POSITIONS, -1, 4)
            Else
                TRADING_TABLE(i, 8) = MAX_SHARES
            End If
        Else
                TRADING_TABLE(i, 8) = ""
        End If
    End If
    
'-------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------

    If TRADING_TABLE(i, 5) = 1 Then ' Long Position
        If TRADING_TABLE(i, 6) <> "" Then
            TRADING_TABLE(i, 9) = (TRADING_TABLE(i, 6) * _
                    TRADING_TABLE(i, 8)) - (TRADING_TABLE(i, 4) * _
                        TRADING_TABLE(i, 8)) - (2 * COST_TRADE)
        Else
            TRADING_TABLE(i, 9) = ""
        
        End If
    Else 'Short or -1
        If TRADING_TABLE(i, 6) <> "" Then
            TRADING_TABLE(i, 9) = (TRADING_TABLE(i, 4) * _
                    TRADING_TABLE(i, 8)) - (TRADING_TABLE(i, 6) * _
                        TRADING_TABLE(i, 8)) - (2 * COST_TRADE)
        Else
            TRADING_TABLE(i, 9) = ""
        End If
    End If

'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------

    If TRADING_TABLE(i, 9) <> "" Then
        TEMP_SUM = TEMP_SUM + TRADING_TABLE(i, 9)
        If TRADING_TABLE(i, 9) > TEMP_MAX Then: TEMP_MAX = TRADING_TABLE(i, 9)
        If TRADING_TABLE(i, 9) < TEMP_MIN Then: TEMP_MIN = TRADING_TABLE(i, 9)
        
        If TRADING_TABLE(i, 9) > 0 Then
            COUNT_WIN = COUNT_WIN + 1
            SUM_WIN = SUM_WIN + TRADING_TABLE(i, 9)
        End If
        If TRADING_TABLE(i, 9) < 0 Then
            COUNT_LOSS = COUNT_LOSS + 1
            SUM_LOSS = SUM_LOSS + TRADING_TABLE(i, 9)
        End If
    End If
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
    
    If TRADING_TABLE(i, 6) <> "" Then
        TRADING_TABLE(i, 10) = TEMP_SUM
    Else
        TRADING_TABLE(i, 10) = ""
    End If

'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------


    If TRADING_TABLE(i, 6) = "" And TRADING_TABLE(i, 4) <> "" Then
          TRADING_TABLE(i, 11) = TRADING_TABLE(i, 4) * TRADING_TABLE(i, 8)
          
          SUM_POSITIONS = SUM_POSITIONS + TRADING_TABLE(i, 11)
          COUNT_POSITIONS = COUNT_POSITIONS + 1
    Else
          TRADING_TABLE(i, 11) = ""
    End If
    
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
    
    If TRADING_TABLE(i, 4) <> "" Then
        If (TRADING_TABLE(i, 5) = 1) Then
            TRADING_TABLE(i, 12) = PROFIT_TARGET * TRADING_TABLE(i, 4) _
                            + TRADING_TABLE(i, 4)
        Else 'Short or -1
            TRADING_TABLE(i, 12) = TRADING_TABLE(i, 4) - PROFIT_TARGET _
                            * TRADING_TABLE(i, 4)
        End If
    Else
            TRADING_TABLE(i, 12) = ""
    End If
Next i

'---------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------

For i = 1 To NROWS
    
'-------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------
    
    If TRADING_TABLE(i, 4) <> "" Then
        If (TRADING_TABLE(i, 5) = 1) Then
            TRADING_TABLE(i, 13) = ((TRADING_TABLE(i, 4) * TRADING_TABLE(i, 8)) - _
                    (((MAX_RISK * BEG_BALANCE) + (TEMP_SUM * MAX_RISK)) / _
                            NO_POSITIONS)) / TRADING_TABLE(i, 8) - TRADING_TABLE(i, 4)
        Else 'Short or -1
            TRADING_TABLE(i, 13) = ((TRADING_TABLE(i, 4) * TRADING_TABLE(i, 8)) + _
                    (((MAX_RISK * BEG_BALANCE) + (TEMP_SUM * MAX_RISK)) / _
                            NO_POSITIONS)) / TRADING_TABLE(i, 8) - TRADING_TABLE(i, 4)
        End If
        TRADING_TABLE(i, 14) = TRADING_TABLE(i, 4) + TRADING_TABLE(i, 13)
    Else
        TRADING_TABLE(i, 13) = ""
        TRADING_TABLE(i, 14) = ""
    End If
    
'-------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------
            
    If i = 1 Then
        If (TRADING_TABLE(i, 5) <> "") And (TRADING_TABLE(i, 4) <> "") Then
            If (TRADING_TABLE(i, 5) = 1) Then
                    TRADING_TABLE(i, 15) = (TRADING_TABLE(i, 4) - _
                    TRADING_TABLE(i, 14)) * -TRADING_TABLE(i, 8) - 2 * COST_TRADE
            Else
                TRADING_TABLE(i, 15) = (TRADING_TABLE(i, 14) - _
                    TRADING_TABLE(i, 4)) * TRADING_TABLE(i, 8) * _
                    -1 - 2 * COST_TRADE
            End If
        Else
            TRADING_TABLE(i, 15) = ""
        End If
    Else
        If (TRADING_TABLE(i, 4) <> "") Then
            If (TRADING_TABLE(i, 5) = 1) Then
                TRADING_TABLE(i, 15) = (TRADING_TABLE(i, 4) - _
                    TRADING_TABLE(i, 14)) * -TRADING_TABLE(i, 8) - 2 * COST_TRADE
            Else
                TRADING_TABLE(i, 15) = (TRADING_TABLE(i, 14) - _
                    TRADING_TABLE(i, 4)) * TRADING_TABLE(i, 8) * _
                    -1 - 2 * COST_TRADE
            End If
        Else
            TRADING_TABLE(i, 15) = ""
        End If
    End If
    
'-------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------
    
    If ((TRADING_TABLE(i, 2) <> "") And (TRADING_TABLE(i, 16) <> "")) Then
          TRADING_TABLE(i, 17) = NETWORKDAYS_FUNC(TRADING_TABLE(i, 2), _
                     TRADING_TABLE(i, 16), HOLIDAYS_RNG)
    Else
          TRADING_TABLE(i, 17) = ""
    End If
'---------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------
Next i
'---------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------
        
    Select Case OUTPUT
        Case 0
        ReDim TRADING_SUMMARY(1 To 7, 1 To 8)
        
        TRADING_SUMMARY(1, 1) = "MAX HIGH" 'Best trade profit.
        TRADING_SUMMARY(2, 1) = "MAX LOW" 'Worst trade profit.
        TRADING_SUMMARY(3, 1) = "NO WINS" 'Number of Winning Trades
        TRADING_SUMMARY(4, 1) = "NO LOSSES" 'Number of Losing Trades
        TRADING_SUMMARY(5, 1) = "PROFIT TARGET" 'This is the value
        'used to calculate the exit target. By default, it is 8%.
        TRADING_SUMMARY(6, 1) = "POSITIONS" 'This value is changed for
        'each trade to reflect how many trades there are for the day.
        'Remember, you MUST hardcode the shares bought for each position
        'or the number of shares for closed positions will be recalculated,
        'giving erroneous results. Change this value to the number of
        'positions entered each day.
        TRADING_SUMMARY(7, 1) = "MAX RISK" 'This is a number you enter -
        'it tells you where to place a trailing stop to limit your
        'risk to your tolerance level.
    
        TRADING_SUMMARY(1, 2) = TEMP_MAX
        TRADING_SUMMARY(2, 2) = TEMP_MIN
        TRADING_SUMMARY(3, 2) = COUNT_WIN
        TRADING_SUMMARY(4, 2) = COUNT_LOSS
        TRADING_SUMMARY(5, 2) = PROFIT_TARGET
        TRADING_SUMMARY(6, 2) = NO_POSITIONS
        TRADING_SUMMARY(7, 2) = MAX_RISK
        
        TRADING_SUMMARY(1, 3) = "AVG WIN"
        TRADING_SUMMARY(2, 3) = "AVG LOSS"
        TRADING_SUMMARY(3, 3) = "MAX SHARES" 'The maximum number of shares
        'you will have in any position. You can change this number to
        'reflect a more conservative approach. Type in 100 for the maximum
        'number of shares and the amount of your investment will be
        'significantly reduced. NOTE: This value is used only for initial
        'calculation. Once the shares field is hardcoded, the shares field
        'will not be updated.

        TRADING_SUMMARY(4, 3) = "POSITIONS OPEN" 'The actual number of open positions.
        'Keep a close eye on this as you don't get too deeply into margin!
        TRADING_SUMMARY(5, 3) = "DDLR RATIO" 'Discipline, Direction, Leverage
        'and Risk ratio. Should be above 3 to maintain confidence in
        'using margin.
        TRADING_SUMMARY(6, 3) = "W/L RATIO" 'The win/loss ratio.
        TRADING_SUMMARY(7, 3) = ""
    
        If COUNT_WIN > 0 Then
            If TRADING_SUMMARY(2, 2) <> upper_limit Then
                  TRADING_SUMMARY(1, 4) = SUM_WIN / COUNT_WIN
            Else
                  TRADING_SUMMARY(1, 4) = ""
            End If
        Else
            TRADING_SUMMARY(1, 4) = ""
        End If

        If COUNT_LOSS > 0 Then
            If TRADING_SUMMARY(2, 2) <> upper_limit Then
                  TRADING_SUMMARY(2, 4) = SUM_LOSS / COUNT_LOSS
            Else
                  TRADING_SUMMARY(2, 4) = ""
            End If
        Else
            TRADING_SUMMARY(2, 4) = ""
        End If
                
        TRADING_SUMMARY(3, 4) = MAX_SHARES
        TRADING_SUMMARY(4, 4) = COUNT_POSITIONS
        
        If (TRADING_SUMMARY(1, 4) <> "") And (TRADING_SUMMARY(2, 4) <> "") Then
              TRADING_SUMMARY(5, 4) = Abs((TRADING_SUMMARY(3, 2) / _
                                        TRADING_SUMMARY(4, 2)) * _
                                        (TRADING_SUMMARY(1, 4) / TRADING_SUMMARY(2, 4)))
        Else
              TRADING_SUMMARY(5, 4) = ""
        End If
        
        If TRADING_SUMMARY(4, 2) <> 0 Then
              TRADING_SUMMARY(6, 4) = TRADING_SUMMARY(3, 2) / TRADING_SUMMARY(4, 2)
        Else
              TRADING_SUMMARY(6, 4) = TRADING_SUMMARY(3, 2)
        End If
        
        TRADING_SUMMARY(7, 4) = ""


        
        TRADING_SUMMARY(1, 5) = "MARGIN %" 'With a 'normal' account, the maximum
        'that margin can be is 100%. Keep in mind that as a Pattern Day Trader,
        'this value can be as high as 400%. However, with most brokerages, you
        'CAN NOT hold a position overnight with more than 100% of margin used
        'or you will get a Reg T Margin Call in the morning.
        
        TRADING_SUMMARY(2, 5) = "PERFORMANCE/DAY"
        TRADING_SUMMARY(3, 5) = "DAYS SYSTEM" 'This one returns the number of
        'actual trading days the system has been in effect.
        TRADING_SUMMARY(4, 5) = "TOTAL MARGIN" 'The current total margin
        'available - as a function of the value you have entered for Margin %.
        TRADING_SUMMARY(5, 5) = "PERFORMANCE" 'The current performance
        'of the account.
        TRADING_SUMMARY(6, 5) = "COST/TRADE" 'This is the one way cost per trade.
        'Remember, there is this cost for entering the trade and the same cost
        'again for exiting the trade. If you use IB or Cybertrader, you will
        'need to calculate this differently as a price per function of shares
        'purchased.
        
        TRADING_SUMMARY(7, 5) = ""
    
    
        TRADING_SUMMARY(1, 6) = MARGIN_PERCENTAGE
        
        If DAY_MAX > DAY_MIN Then
            TRADING_SUMMARY(3, 6) = NETWORKDAYS_FUNC(DAY_MIN, DAY_MAX, HOLIDAYS_RNG)
        Else
            TRADING_SUMMARY(3, 6) = ""
        End If

        
        TRADING_SUMMARY(4, 6) = TEMP_SUM * MARGIN_PERCENTAGE + _
                    (BEG_BALANCE * MARGIN_PERCENTAGE) + TEMP_SUM
        
        If BEG_BALANCE <> 0 Then
            TRADING_SUMMARY(5, 6) = TEMP_SUM / BEG_BALANCE
        Else
            TRADING_SUMMARY(5, 6) = ""
        End If

        TRADING_SUMMARY(6, 6) = COST_TRADE
        TRADING_SUMMARY(7, 6) = ""


        If TRADING_SUMMARY(3, 6) <> "" And TRADING_SUMMARY(5, 6) <> "" Then
            If TRADING_SUMMARY(3, 6) <> 0 Then
                TRADING_SUMMARY(2, 6) = TRADING_SUMMARY(5, 6) / TRADING_SUMMARY(3, 6)
            Else
                TRADING_SUMMARY(2, 6) = ""
            End If
        Else
                TRADING_SUMMARY(2, 6) = ""
        End If

        TRADING_SUMMARY(1, 7) = "BEGINNING BALANCE" 'Your beginning balance
        TRADING_SUMMARY(2, 7) = "INITIAL MARGIN"
        TRADING_SUMMARY(3, 7) = "ROI ANNUALIZED"
        TRADING_SUMMARY(4, 7) = "RETURN $" 'The amount you have currently
        'made or lost on the account.
        TRADING_SUMMARY(5, 7) = "ACCOUNT VALUE"
        TRADING_SUMMARY(6, 7) = "AMOUNT INVESTED"
        TRADING_SUMMARY(7, 7) = ""
        
        TRADING_SUMMARY(1, 8) = BEG_BALANCE
        TRADING_SUMMARY(2, 8) = BEG_BALANCE * MARGIN_PERCENTAGE
        
        If TRADING_SUMMARY(2, 6) <> "" Then
            TRADING_SUMMARY(3, 8) = TRADING_SUMMARY(2, 6) * COUNT_BASIS
        Else
            TRADING_SUMMARY(3, 8) = ""
        End If
            
        TRADING_SUMMARY(4, 8) = TEMP_SUM
        
        TRADING_SUMMARY(5, 8) = BEG_BALANCE + TEMP_SUM
        TRADING_SUMMARY(6, 8) = SUM_POSITIONS
        TRADING_SUMMARY(7, 8) = ""
    
        PORT_TRACK_TRADING_ACTIVITIES_FUNC = TRADING_SUMMARY
        
        Case Else
        PORT_TRACK_TRADING_ACTIVITIES_FUNC = TRADING_TABLE
    End Select
        

Exit Function
ERROR_LABEL:
PORT_TRACK_TRADING_ACTIVITIES_FUNC = Err.number
End Function


Sub MGMT_TRADING_ACTIVITIES_FUNC()
Dim NO_ASSETS As Variant
NO_ASSETS = Application.InputBox("No Assets", "Trading Activities")
If IsNumeric(NO_ASSETS) = False Then: Exit Sub

Call EXCEL_TURN_OFF_EVENTS_FUNC
Call RNG_PORT_TRACK_TRADING_ACTIVITIES_FUNC(NO_ASSETS, ActiveWorkbook)
Call EXCEL_TURN_ON_EVENTS_FUNC

End Sub

Function RNG_PORT_TRACK_TRADING_ACTIVITIES_FUNC(Optional ByVal NO_ASSETS As Long = 100, _
Optional ByRef SRC_WBOOK As Excel.Workbook)

Dim j As Long

Dim TEMP_STR As String
Dim HEADINGS_ARR As Variant
Dim RC_SWITCH_FLAG As Boolean
Dim TEMP_RNG As Excel.Range
Dim DATA_RNG As Excel.Range
Dim DST_RNG As Excel.Range
Dim DST_WSHEET As Excel.Worksheet

On Error GoTo ERROR_LABEL

If SRC_WBOOK Is Nothing Then: Set SRC_WBOOK = ActiveWorkbook
Set DST_WSHEET = WSHEET_ADD_FUNC(PARSE_CURRENT_TIME_FUNC("_"), SRC_WBOOK)
ActiveWindow.DisplayGridlines = False
GoSub FORMAT_WSHEET_LINE

Set DST_RNG = DST_WSHEET.Cells(2, 2)

'-------------------------------------------------------------------------------------------------------------
With DST_RNG
'-------------------------------------------------------------------------------------------------------------
    .Cells(1, 1).value = "BEGINNING BALANCE"
    .Cells(1, 1).Font.Bold = True
    TEMP_STR = "Your beginning balance."
    .Cells(1, 1).AddComment TEMP_STR
    .Cells(1, 2).value = 10000
    .Cells(1, 2).Font.ColorIndex = 5
    .Cells(1, 2).NumberFormat = "#,##0.00"
    
    .Cells(2, 1).value = "PROFIT TARGET"
    .Cells(2, 1).Font.Bold = True
    TEMP_STR = "This is the value used to calculate the exit target. By default, it is 8%."
    .Cells(2, 1).AddComment TEMP_STR
    .Cells(2, 2).value = 0.08
    .Cells(2, 2).Font.ColorIndex = 5
    .Cells(2, 2).NumberFormat = "0.00%"
    
    .Cells(3, 1).value = "NO POSITIONS"
    .Cells(3, 1).Font.Bold = True
    TEMP_STR = "This value is changed for each trade to reflect how many trades there are for the day." & _
    " Remember, you MUST hardcode the shares bought for each position or the number of shares for " & _
    "closed positions will be recalculated, giving erroneous results. Change this value to the number " & _
    "of positions entered each day."
    .Cells(3, 1).AddComment TEMP_STR
    .Cells(3, 2).value = 10
    .Cells(3, 2).Font.ColorIndex = 5
    .Cells(3, 2).NumberFormat = "#,##0"
    
    .Cells(4, 1).value = "MAX RISK"
    .Cells(4, 1).Font.Bold = True
    TEMP_STR = "This is a number you enter - it tells you where to place a trailing stop to limit " & _
    "your risk to your tolerance level."
    .Cells(4, 1).AddComment TEMP_STR
    .Cells(4, 2).value = 0.03
    .Cells(4, 2).Font.ColorIndex = 5
    .Cells(4, 2).NumberFormat = "0.00%"
    
    .Cells(5, 1).value = "MAX SHARES"
    .Cells(5, 1).Font.Bold = True
    TEMP_STR = "The maximum number of shares you will have in any position. You can change this number" & _
    "to reflect a more conservative approach. Type in 100 for the maximum number of shares and the" & _
    "amount of your investment will be significantly reduced. NOTE: This value is used only for initial " & _
    "calculation. Once the shares field is hardcoded, the shares field will not be updated."
    .Cells(5, 1).AddComment TEMP_STR
    .Cells(5, 2).value = 1000
    .Cells(5, 2).Font.ColorIndex = 5
    .Cells(5, 2).NumberFormat = "#,##0"
    
    .Cells(6, 1).value = "MARGIN PERCENTAGE"
    .Cells(6, 1).Font.Bold = True
    TEMP_STR = "With a 'normal' account, the maximum that margin can be is 100%. Keep in" & _
    "mind that as a Pattern Day Trader, this value can be as high as 400%. However," & _
    "with most brokerages, you CAN NOT hold a position overnight with more than 100% of" & _
    "margin used or you will get a Reg T Margin Call in the morning."
    .Cells(6, 1).AddComment TEMP_STR
    .Cells(6, 2).value = 0.01
    .Cells(6, 2).Font.ColorIndex = 5
    .Cells(6, 2).NumberFormat = "0.00%"
    
    .Cells(7, 1).value = "COST/TRADE"
    .Cells(7, 1).Font.Bold = True
    TEMP_STR = "This is the one way cost per trade. Remember, there is this cost for" & _
    "entering the trade and the same cost again for exiting the trade. If you use IB or" & _
    "Cybertrader, you will need to calculate this differently as a price per function" & _
    "of shares purchased."
    .Cells(7, 1).AddComment TEMP_STR
    .Cells(7, 2).value = 4
    .Cells(7, 2).Font.ColorIndex = 5
    .Cells(7, 2).NumberFormat = "#,##0.00"
    
    .Cells(8, 1).value = "COUNT BASIS"
    .Cells(8, 1).Font.Bold = True
    TEMP_STR = "Trading days per year"
    .Cells(8, 1).AddComment TEMP_STR
    .Cells(8, 2).value = 251
    .Cells(8, 2).Font.ColorIndex = 5
    .Cells(8, 2).NumberFormat = "#,##0"
    
    RC_SWITCH_FLAG = False
    Set TEMP_RNG = Range(.Cells(1, 1), .Cells(8, 2))
    GoSub FORMAT_BOX_LINE
    TEMP_RNG.Columns(1).InsertIndent 1
    
'-------------------------------------------------------------------------------------------------------------
    HEADINGS_ARR = Array("SYMBOL", "SETTLEMENT", "PREVIOUS CLOSE", "PRICE IN", _
                  "POSITION", "SHARES", "DAY OUT", "PRICE OUT")
    
    Set TEMP_RNG = Range(.Cells(22, 19), .Cells(22 + NO_ASSETS, 26))
    With TEMP_RNG.Rows(1)
        .value = HEADINGS_ARR
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        TEMP_STR = "Long(1)/Short(-1)"
        .Cells(1, 5).AddComment TEMP_STR
    End With
    RC_SWITCH_FLAG = True
    GoSub FORMAT_BOX_LINE
    Set DATA_RNG = Range(.Cells(23, 19), .Cells(22 + NO_ASSETS, 26))
    With DATA_RNG
        .Font.ColorIndex = 5
        .Columns(1).HorizontalAlignment = xlCenter
        .Columns(2).NumberFormat = "[$-409]d-mmm-yy;@"
        .Columns(3).NumberFormat = "#,##0.00"
        .Columns(4).NumberFormat = "#,##0.00"
        With .Columns(5)
            .value = 1
            .Font.ColorIndex = 3
            .HorizontalAlignment = xlCenter
            .NumberFormat = "0"
            With .Validation
                .Delete
                .Add Type:=xlValidateList, _
                AlertStyle:=xlValidAlertStop, _
                Operator:=xlBetween, _
                Formula1:="1,-1"
            End With
        End With
        .Columns(6).NumberFormat = "#,##0"
        .Columns(7).NumberFormat = "[$-409]d-mmm-yy;@"
        .Columns(8).NumberFormat = "#,##0.00"
    End With
'-------------------------------------------------------------------------------------------------------------
    RC_SWITCH_FLAG = False
    .Cells(10, 1).value = "CONTROL"
    .Cells(10, 1).Font.Bold = True
    With .Cells(10, 2)
        .value = False
        .HorizontalAlignment = xlCenter
        .Font.ColorIndex = 3
        With .Validation
            .Delete
            .Add Type:=xlValidateList, _
            AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, _
            Formula1:="TRUE,FALSE"
        End With
    End With
    Set TEMP_RNG = Range(.Cells(10, 1), .Cells(10, 2))
    GoSub FORMAT_BOX_LINE
    TEMP_RNG.Columns(1).InsertIndent 1
    
    Set TEMP_RNG = Range(.Cells(12, 1), .Cells(18, 8))
    GoSub FORMAT_BOX_LINE
    With TEMP_RNG.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    TEMP_RNG.FormulaArray = _
            "=IF(" & DST_RNG.Cells(10, 2).Address & "=FALSE," & """" & "-" & """" & "," & _
                "PORT_TRACK_TRADING_ACTIVITIES_FUNC(" & _
                DATA_RNG.Columns(1).Address & "," & _
                DATA_RNG.Columns(2).Address & "," & _
                DATA_RNG.Columns(3).Address & "," & _
                DATA_RNG.Columns(4).Address & "," & _
                DATA_RNG.Columns(5).Address & "," & _
                DATA_RNG.Columns(6).Address & "," & _
                DATA_RNG.Columns(7).Address & "," & _
                DATA_RNG.Columns(8).Address & "," & _
                .Cells(1, 2).Address & "," & _
                .Cells(2, 2).Address & "," & _
                .Cells(3, 2).Address & "," & _
                .Cells(4, 2).Address & "," & _
                .Cells(5, 2).Address & "," & _
                .Cells(6, 2).Address & "," & _
                .Cells(7, 2).Address & "," & _
                .Cells(8, 2).Address & ",0))"
    
    With TEMP_RNG
        TEMP_STR = "Best trade profit."
        .Cells(1, 1).AddComment TEMP_STR
        
        TEMP_STR = "Worst trade profit."
        .Cells(2, 1).AddComment TEMP_STR
        
        TEMP_STR = "Number of winning trades."
        .Cells(3, 1).AddComment TEMP_STR
        
        TEMP_STR = "Number of losing trades."
        .Cells(4, 1).AddComment TEMP_STR
        
        TEMP_STR = "The actual number of open positions. Keep a close eye " & _
        "on this as you don't get too deeply into margin!"
        .Cells(4, 3).AddComment TEMP_STR
        
        TEMP_STR = "Discipline, Direction, Leverage and Risk ratio. Should be " & _
        "above 3 to maintain confidence in using margin."
        .Cells(5, 3).AddComment TEMP_STR
        
        TEMP_STR = "Winn Loss Ratio"
        .Cells(6, 3).AddComment TEMP_STR
        
        TEMP_STR = "Number of actual trading days the system has been in effect. " & _
        "You need to build your own holiday list."
        .Cells(3, 5).AddComment TEMP_STR
        
        TEMP_STR = "The current total margin available - as a function of the value " & _
        "you have entered for Margin %."
        .Cells(4, 5).AddComment TEMP_STR
        
        TEMP_STR = "Current performance of the account"
        .Cells(5, 5).AddComment TEMP_STR
        
        TEMP_STR = "The amount you have currently made or lost on the account."
        .Cells(4, 7).AddComment TEMP_STR
        
        Union(.Cells(1, 2), .Cells(2, 2), .Cells(1, 4), .Cells(2, 4), _
              .Cells(5, 4), .Cells(6, 4), .Cells(4, 6), .Cells(6, 6), _
              .Cells(1, 8), .Cells(2, 8), .Cells(4, 8), .Cells(5, 8), _
              .Cells(6, 8)).NumberFormat = "#,##0.00"
              
        Union(.Cells(3, 2), .Cells(4, 2), .Cells(6, 2), .Cells(4, 4), _
              .Cells(3, 6), .Cells(3, 6)).NumberFormat = "#,##0"
              
        Union(.Cells(5, 2), .Cells(7, 2), .Cells(1, 6), .Cells(2, 6), _
              .Cells(5, 6), .Cells(3, 8)).NumberFormat = "0.00%"
    End With
    
    For j = 1 To 8 Step 2
        Set TEMP_RNG = Range(.Cells(12, j), .Cells(18, j))
        TEMP_RNG.Font.Bold = True
        GoSub FORMAT_BOX_LINE
        TEMP_RNG.Columns(1).InsertIndent 1
    Next j
'-------------------------------------------------------------------------------------------------------------
    RC_SWITCH_FLAG = False
    .Cells(20, 1).value = "CONTROL"
    .Cells(20, 1).Font.Bold = True
    With .Cells(20, 2)
        .value = False
        .HorizontalAlignment = xlCenter
        .Font.ColorIndex = 3
        With .Validation
            .Delete
            .Add Type:=xlValidateList, _
            AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, _
            Formula1:="TRUE,FALSE"
        End With
    End With
    Set TEMP_RNG = Range(.Cells(20, 1), .Cells(20, 2))
    GoSub FORMAT_BOX_LINE
    TEMP_RNG.Columns(1).InsertIndent 1

    Set TEMP_RNG = Range(.Cells(22, 1), .Cells(22 + NO_ASSETS, 17))
    TEMP_RNG.FormulaArray = _
            "=IF(" & DST_RNG.Cells(20, 2).Address & "=FALSE," & """" & "-" & """" & "," & _
                "PORT_TRACK_TRADING_ACTIVITIES_FUNC(" & _
                DATA_RNG.Columns(1).Address & "," & _
                DATA_RNG.Columns(2).Address & "," & _
                DATA_RNG.Columns(3).Address & "," & _
                DATA_RNG.Columns(4).Address & "," & _
                DATA_RNG.Columns(5).Address & "," & _
                DATA_RNG.Columns(6).Address & "," & _
                DATA_RNG.Columns(7).Address & "," & _
                DATA_RNG.Columns(8).Address & "," & _
                .Cells(1, 2).Address & "," & _
                .Cells(2, 2).Address & "," & _
                .Cells(3, 2).Address & "," & _
                .Cells(4, 2).Address & "," & _
                .Cells(5, 2).Address & "," & _
                .Cells(6, 2).Address & "," & _
                .Cells(7, 2).Address & "," & _
                .Cells(8, 2).Address & ",1))"
    
    With TEMP_RNG
        .NumberFormat = "#,##0.00"
        Union(.Columns(1), .Columns(5)).HorizontalAlignment = xlCenter
        Union(.Columns(2), .Columns(16)).NumberFormat = "[$-409]d-mmm-yy;@"
        Union(.Columns(5), .Columns(8), .Columns(17)).NumberFormat = "#,##0"
    End With
    
    RC_SWITCH_FLAG = True
    GoSub FORMAT_BOX_LINE
    With TEMP_RNG.Rows(1)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    
        TEMP_STR = "Only used for calculating position size for upcoming trade."
        .Cells(1, 3).AddComment TEMP_STR
        
        TEMP_STR = "Completing this field (at time of entry) calculates the target value."
        .Cells(1, 4).AddComment TEMP_STR
    
        TEMP_STR = "The price of the stock  when you sold it, or the last price (for " & _
        "tracking P/L). You must manually enter this value. "
        .Cells(1, 6).AddComment TEMP_STR
        
        TEMP_STR = "The percentage gain or loss for the trade. Note that the cost " & _
        "of commissions is NOT figured in this result This is the %G/L for the trade.."
        .Cells(1, 7).AddComment TEMP_STR
        
        TEMP_STR = "Once this field is calculated, it MUST be hardcoded to the calculated " & _
        "value. Otherwise changes in previous, open trades will cause this value to be recalculated."
        .Cells(1, 8).AddComment TEMP_STR
        
        TEMP_STR = "The profit or loss of the trade. Note that commission costs are included " & _
        "in this result."
        .Cells(1, 9).AddComment TEMP_STR
        
        TEMP_STR = "This is a running total of the current result of your trading."
        .Cells(1, 10).AddComment TEMP_STR
        
        TEMP_STR = "This field is populated only during the time the trade (or any portion " & _
        "of) is active."
        .Cells(1, 11).AddComment TEMP_STR
        
        TEMP_STR = "Target is determined by the value in Profit Target above."
        .Cells(1, 12).AddComment TEMP_STR
        
        TEMP_STR = "The 'suggested' trailing stop, as a function of your value for risk."
        .Cells(1, 13).AddComment TEMP_STR
        
        TEMP_STR = "What you should set a fixed stop at as a function of your value for risk."
        .Cells(1, 14).AddComment TEMP_STR
        
        TEMP_STR = "This is how much you are risking, not including any slippage. Includes " & _
        "trade cost. Populated at Price In time."
        .Cells(1, 15).AddComment TEMP_STR
        
        TEMP_STR = "Day the trade is closed. Fill this in manually."
        .Cells(1, 16).AddComment TEMP_STR
        
        TEMP_STR = "Days the trade lived. Useful if you want to track your trade length."
        .Cells(1, 17).AddComment TEMP_STR
    
    End With

'-------------------------------------------------------------------------------------------------------------
    
End With

Exit Function
'------------------------------------------------------------------------------
FORMAT_WSHEET_LINE:
'------------------------------------------------------------------------------
With DST_WSHEET.Cells
    With .Font
        .name = "Courier New"
        .Size = 10
    End With
    .VerticalAlignment = xlCenter
    .RowHeight = 15
    .ColumnWidth = 20
    .Columns(1).ColumnWidth = 3
End With
Return
'-------------------------------------------------------------------------------------
FORMAT_BOX_LINE:
'-------------------------------------------------------------------------------------
With TEMP_RNG
    .Borders(xlDiagonalDown).LineStyle = xlNone
    .Borders(xlDiagonalUp).LineStyle = xlNone
    With .Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .WEIGHT = xlThin
        .ColorIndex = xlAutomatic
    End With
    With .Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .WEIGHT = xlThin
        .ColorIndex = xlAutomatic
    End With
    With .Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .WEIGHT = xlThin
        .ColorIndex = xlAutomatic
    End With
    With .Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .WEIGHT = xlThin
        .ColorIndex = xlAutomatic
    End With
    With .Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .WEIGHT = xlThin
        .ColorIndex = xlAutomatic
    End With
    With .Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .WEIGHT = xlThin
        .ColorIndex = xlAutomatic
    End With
    With IIf(RC_SWITCH_FLAG = True, .Rows(1).Interior, .Columns(1).Interior)
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End With
'-------------------------------------------------------------------------------------
Return
'-------------------------------------------------------------------------------------

ERROR_LABEL:
RNG_PORT_TRACK_TRADING_ACTIVITIES_FUNC = Err.number
End Function
