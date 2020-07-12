Attribute VB_Name = "FINAN_ASSET_SYSTEM_PREMARK_LIBR"
'-------------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------------
Option Explicit
Option Base 1
'-------------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------------

'preMARKET Activity ... and (maybe) an investment strategy
'I was thinking about investor over-reaction to good (or bad) news.
'A penny less in reported earnings and the stock may drop like a rock ...
'an over-reaction, I think. Then, later that same day, it recovers.
'Always? No, of course not. But suppose we know that a stock will open
'UP by, say, 2% from yesterday's close. Up 2%? Isn't that a large jump?
'Yes, like a 200 jump in a 10,000 DOW. Probably an over-reaction, so we might
'expect the stock to drop later that same day.

'For example, ATYT is a very volatile stock so we might expect that large changes
'might occur frequently (from day to day). Very volatile? Yes. The stock price gapped
'up at the open on Nov 3, 2004 ... but dwindled during the day.

'Okay, let's see if we could have made any money by selling at the open when the stock
'price is up dramatically, then buying later in the day ... say at the close. So here's
'what we'll do ...

'1. We start with 1000 shares of stock and $3,000 in Cash.
'2. We sell all 1000 shares at the open when the stock increases by 2% from yesterday's close.
'3. We always buy back our 1000 shares at the close (regardless of price).
'4. If the stock price falls from open to close, we made money which we stick into Cash.
'5. If the stock price rises from open to close, we need to take some money from Cash to buy back
'   the 1000 shares (because of the higher closing price).

'Reference: http://www.gummy-stuff.org/preMkt.htm

Function ASSET_PRE_MARKET_SIGNAL_FUNC(ByVal TICKER_STR As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal END_DATE As Date, _
Optional ByVal BUY_PERCENT As Double = -0.02, _
Optional ByVal SELL_PERCENT As Double = 0.02, _
Optional ByVal SYSTEM_SHARES As Long = 1000, _
Optional ByVal COST_PER_TRADE As Double = 15, _
Optional ByVal INITIAL_CASH As Double = 0, _
Optional ByVal CASH_RATE As Double = 0.02, _
Optional ByVal COUNT_BASIS As Double = 365, _
Optional ByVal OUTPUT As Integer = 0)

Dim i As Long
Dim j As Long
Dim NROWS As Long

Dim MEAN_VAL As Double
Dim VOLAT_VAL As Double

Dim TEMP_FACTOR As Double
Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

If IsArray(TICKER_STR) = True Then
    DATA_MATRIX = TICKER_STR
Else
    DATA_MATRIX = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, _
                  START_DATE, END_DATE, "DAILY", _
                  "DOHLCVA", False, True, True)
End If
NROWS = UBound(DATA_MATRIX, 1)


TEMP_FACTOR = (1 + CASH_RATE) ^ (1 / COUNT_BASIS)
ReDim TEMP_MATRIX(0 To NROWS, 1 To 17)

TEMP_MATRIX(0, 1) = "DATE"
TEMP_MATRIX(0, 2) = "OPEN"
TEMP_MATRIX(0, 3) = "HIGH"
TEMP_MATRIX(0, 4) = "LOW"
TEMP_MATRIX(0, 5) = "CLOSE"
TEMP_MATRIX(0, 6) = "VOLUME"
TEMP_MATRIX(0, 7) = "ADJ. CLOSE"

TEMP_MATRIX(0, 8) = "OPEN/CLOSE"

TEMP_MATRIX(0, 9) = "BUY WHEN DN BY " & Format(Abs(BUY_PERCENT), "0.00%")
TEMP_MATRIX(0, 10) = "SELL WHEN UP BY " & Format(SELL_PERCENT, "0.00%")

TEMP_MATRIX(0, 11) = "CASH"
TEMP_MATRIX(0, 12) = Format(SYSTEM_SHARES, "0") & " SYSTEM SHARES"
TEMP_MATRIX(0, 13) = "CASH BUY"
TEMP_MATRIX(0, 14) = "CASH SELL"
TEMP_MATRIX(0, 15) = "SYSTEM BALANCE"
TEMP_MATRIX(0, 16) = "BUY THEN SELL"
TEMP_MATRIX(0, 17) = "SELL THEN BUY"

i = 1
For j = 1 To 7: TEMP_MATRIX(i, j) = DATA_MATRIX(i, j): Next j
TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 6) / 1000
For j = 8 To 17: TEMP_MATRIX(i, j) = "": Next j
TEMP_MATRIX(i, 11) = INITIAL_CASH

MEAN_VAL = 0
For i = 2 To NROWS
    For j = 1 To 7: TEMP_MATRIX(i, j) = DATA_MATRIX(i, j): Next j
    TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 6) / 1000
    TEMP_MATRIX(i, 8) = TEMP_MATRIX(i, 2) / TEMP_MATRIX(i - 1, 7) - 1
    TEMP_MATRIX(i, 9) = IIf(TEMP_MATRIX(i, 8) < BUY_PERCENT, 1, 0)
    TEMP_MATRIX(i, 10) = IIf(TEMP_MATRIX(i, 8) > SELL_PERCENT, 1, 0)
    TEMP_MATRIX(i, 13) = IIf(TEMP_MATRIX(i, 9) = 1, SYSTEM_SHARES * ((TEMP_MATRIX(i, 2) + TEMP_MATRIX(i, 3)) / 2 - TEMP_MATRIX(i, 2)), 0)
    TEMP_MATRIX(i, 14) = IIf(TEMP_MATRIX(i, 10) = 1, SYSTEM_SHARES * (TEMP_MATRIX(i, 2) - (TEMP_MATRIX(i, 2) + TEMP_MATRIX(i, 4)) / 2), 0)
    TEMP_MATRIX(i, 11) = TEMP_MATRIX(i - 1, 11) * TEMP_FACTOR + IIf(TEMP_MATRIX(i, 9) + TEMP_MATRIX(i, 10) = 1, TEMP_MATRIX(i, 13) + TEMP_MATRIX(i, 14) - COST_PER_TRADE, 0)
    TEMP_MATRIX(i, 12) = TEMP_MATRIX(i, 7) * SYSTEM_SHARES
    TEMP_MATRIX(i, 15) = TEMP_MATRIX(i, 11) + TEMP_MATRIX(i, 12)
    If i > 2 Then: MEAN_VAL = MEAN_VAL + (TEMP_MATRIX(i, 15) / TEMP_MATRIX(i - 1, 15) - 1)
    TEMP_MATRIX(i, 16) = IIf(TEMP_MATRIX(i, 9) = 1, TEMP_MATRIX(i, 15), 0)
    TEMP_MATRIX(i, 17) = IIf(TEMP_MATRIX(i, 10) = 1, TEMP_MATRIX(i, 15), 0)
Next i

If OUTPUT = 0 Then
    ASSET_PRE_MARKET_SIGNAL_FUNC = TEMP_MATRIX
    Exit Function
End If

j = NROWS - 2
MEAN_VAL = MEAN_VAL / j

For i = 3 To NROWS
    VOLAT_VAL = VOLAT_VAL + ((TEMP_MATRIX(i, 15) / TEMP_MATRIX(i - 1, 15) - 1) - MEAN_VAL) ^ 2
Next i
VOLAT_VAL = (VOLAT_VAL / j) ^ 0.5

Select Case OUTPUT
Case 1
    ASSET_PRE_MARKET_SIGNAL_FUNC = MEAN_VAL / VOLAT_VAL
Case Else
    ASSET_PRE_MARKET_SIGNAL_FUNC = Array(TEMP_MATRIX, MEAN_VAL / VOLAT_VAL)
End Select

Exit Function
ERROR_LABEL:
ASSET_PRE_MARKET_SIGNAL_FUNC = "--"
End Function

'Note that, at the end of the day, we always have 1000 shares of stock.
'Sometimes we needed cash to pay for the 1000 shares at day's end ... when the price went
'up instead of down. ( That's why we started with some cash, to cover this situation :^)
'However, mostly the closing price was down from the open and we could buy back our 1000
'shares for less than we made on the opening sell ... and we put that extra money into cash.

'Now all you have to do is KNOW when the price will gap up at the open!
'True, but a good indicator is the preMarket activity.
'The idea is to look at the early trading, just before the market opens. If it looks like the
'stock will open significantly higher then (I've found) it does !

'November 5, 2004
'As I write this (Nov 5, 2004) I'm looking at that ATYT stock again
'(love that stock!) ... and I see this
'no , i 'm going to sell at the open ... I think. But I'd like to see the price continue
'to go UP !! As we approach market open, right? Right ... so I wait a bit ...
    
'Then I see this as we get closer to the open (which is 9:30 AM ET):
'It 's goin' down! Yes, but it's still up over 2% so I'll sell 1000 shares at the open ...
'hoping to buy back later in the day. Let me know how you fare, okay?
'Okay ...
'At 9:30 I sold my 1000 shares of ATYT.TO (the Canadian version) at $22.50.
'Yesterday 's close was $22.04. The opening price was up 2.1%
'At 9:40 I put in an order to buy 1000 shares at $22.00 ... and said a small prayer

'You trade ATY.TO but you look at ATYT preMarket activity?
'Uh ... yes, because preMarket activity isn't available for ATY.TO ... ATYT trades
'on the Nasdaq. Shortly after 11 AM I see this
'At 11:35 AM I get my 1000 shares at $22.00. Don't be greedy!
'course, other things can happen ...

'December 13, 2004
'It 's about 9:20 AM on December 13, 2004.
'My brother-in-law phones.
'He says:    "Did you see the news? The CEO of Bombardier resigned."
'I says:     "Whee! The stock will drop like a rock! Bye! Gotta buy some!"
'He says:    "Maybe you should wait for a day or two."
'I says:     "No way! It's the over-reaction to bad news I'm looking at."
'So I quickly login to TD/Waterhouse to buy some (at market), when it opens.
'Aaargh! Bombardier has changed the stock symbol. Why did they have to do that on December 13 ?!
'However, I manage to buy 10,000 shares at $1.91 and ...
'Then you start praying?
'Yes ... but it did close at $2.11 (on Dec 13), so that's a good start, eh?
'Then you sell ... when?
'I kept track the next day, that's Dec 14, and sold the 10,000 shares at $2.30
'And did it go any higher?
'I have no idea, but a 20% gain in 24 hours is good ... right?

'But maybe you could have sold even higher, right?
'don 't be greedy ...
'we 're making the assumption described here ... it works,sometimes

'December 21, 2004
'So ATI Technologies (ATYT on Nasdaq) had an earnings report this morning.
'It looked pretty good to me ... so I checked the preMarket

'So you bought some?
'At the open ... of course ... for $23.75, but the Canadian version, eh?
'Then you start praying?
'Of course.
'So, if that earnings report was good, why did the stock go down?
'Does it matter?
'And you sold later in the day, right?
'Uh ... not exactly.
'Open = $23.75, Close = $23.40
'Now you wait?
'Yes ...
'Open = $23.60, Close = $23.98
'... and wait ...
'Open = $23.92, Close = $23.71
'Waiting for Santa Claus?
'Yes ...


'we 're relying upon investor over-reaction to good or bad news and

'Yeah, but it rarely works, right?
'Well, were making two assumptions   ... which may (or may not) be valid
'We Buy when the stock opens DOWN
'... and we Sell later that day:

'We Sell when the stock opens UP
'... and we Buy later that day:

'Okay, so sometimes it works and sometimes it don't. What does that prove?
'Besides, you sell at the open and you want to buy back at the low
'... but how can you buy at the low when you don't know when you're at the low
'... and if you buy at the open because the price dropped, how can you sell at the
'high when ...?

'Okay, let's think about this ... to see how this strategy would have fared in the past.
'Let's assume when we Buy at the Open we Sell later that day, half-way between the
'Open and the High.

'And when you Sell at the Open you Buy ...?
'Half-way between the Open and the Low. That's the scheme we'll study. It's an approximation
'to what we might do in practice.

'You're telling me!
'Well , It 's something like we did before, remember?
'we didn 't get the Low, but somewhere between the Open and the Low.

'Anyway, let's see what would have happened if we had done the following:
'We use the preMarket activity to predict the Opening price. That'll tell us whether we're
'going to Buy or Sell.
'We start with N = 1000 shares of some (volatile) stock.
'When the stock is UP by 2% (from the previous close, we Sell our N shares.
'We always buy back these N shares at a price mid-way between the Open and the Low.
'(This prescription is because we don't know when the Low will occur.)
'When the stock is DOWN by 2% (from the previous close), we Buy N shares.
'We always sell these N shares at a price mid-way between the Open and the High.
'(This prescription is because we don't know when the High will occur.)
'Since we're in-and-out of the stock on the same day (or out-then-in),
'we 'll assume that our broker doesn't won't require any money for the two transactions
'... except for transaction fees, of course.

'Well, my broker gives me a day or two to pay for purchases. If I BUY then SELL the same day,
'I get charged transaction fees only.
'Note:
'We always own just 1000 shares when the market closes.
'We always Buy and/or Sell 1000 shares when we trade.
'If there are several Buys in a row, we don't have to come up with the money since we Sell
'the same day.

'Money made during a trading day goes into Cash.
'Each day that we trade there are two transaction fees.

'Anyway , Here 's what would have happened over the past year (assuming a trading fee of $15):
'Look at the right-most chart. That's ATI Technologies (ATYT). It's was up by 31% for the year ...
'and volatile!.

'Now look at the left-most chart. Each day we traded, the money we made goes into Cash. That's a
'picture of our Cash holdings.

'Now look at the middle chart. Remember that at the end of the day we always have our 1000 shares
'of stock. The value of that 1000 shares will go up and down just like the stock itself. (See the
'right-most chart?)

'Our total portfolio, however, includes the Cash ... and that's what's plotted in the middle chart.
'It would have been up by 67%, for the year.

'So this past year you made 67% on your investments?
'Uh ... well, no. I'm the world's worst investor ... this is just fun.
'You invent some scheme, look back to see how it would have done, then pretend you actually did
'it that way.

'There were 20 SELL-then-BUY days and 12 BUY-then-SELL days (out of 250 market days for
'the year).

'But why keep all your winnings in Cash? That's a lousy investment, eh?
'It 's safe, risk-free. The market can crash and we get to keep all that Cash. It's ...
'It's still a lousy investment.
'you don 't like the 67% gain?

'Yeah, but this is fiction. Can you give me an up close and personal?
'Here 's a picture of some trades in April, 2004

'That's one confusing chart !!

'Note that we're assuming that when we BUY at the Open, we SELL half-way between the Open and the High.
'And when we SELL at the Open, we BUY half-way between ...

'Can we stop here? This is fiction. How do you know when you're half-way between ...?
'you don 't. We're just testing this strategy, assuming you can't get the High or Low for the day, but
'somewhere in between.

'Can we stop here?
'I should point out that this strategy is lousy if you're investing in something like the S&P 500.
'You'll need to find a volatile stock and ...

