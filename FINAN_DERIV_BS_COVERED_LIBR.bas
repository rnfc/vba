Attribute VB_Name = "FINAN_DERIV_BS_COVERED_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : COVERED_CALL_RETURN_FUNC
'DESCRIPTION   : You buy stock and sell a Call Option, borrowing money to
'complete the transaction. When the Option is exercised you sell at the Strike
'Price and repay the loan. Question: What is your Annualized Return?
'Net Out-of-pocket cost  = Po - A - BS must be positive !!
'LIBRARY       : DERIVATIVES
'GROUP         : RETURN
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function COVERED_CALL_RETURN_FUNC(ByVal AMOUNT_BORROWED As Double, _
ByVal BORROWING_RATE As Double, _
ByVal ACTUAL_PREMIUM As Double, _
ByVal SPOT As Double, _
ByVal STRIKE As Double, _
ByVal CALL_EXERCISED_TENOR As Double, _
ByVal EXPIRATION As Double, _
ByVal RATE As Double, _
ByVal CARRY_COST As Double, _
ByVal SIGMA As Double)

'You buy stock and sell a Call Option, borrowing money to complete
'the transaction. When the Option is exercised you sell at the Strike
'Price and repay the loan.

'Call Option Strike Price:  K =  $30.00
'Current Stock Price:  Po =  $38.00
'Time to Expiry: To = 0.950 years, or 347 days.
'Annual Stock Volatility:  V = 25.00% Note: All this stuff is so we can
'calculate the Black-Scholes premium. You can ignore these if you wish.
'Risk-free Rate:  Rf =   4.00%
'$9.64    This is the "estimated" Black-Scholes premium
'"Actual" Option Premium: BS = $9.30 This is the "actual" premium you receive.
        
'Amount borrowed: A =    $18.00   This is the amount you borrow, originally
'Annual Borrowing Rate: I =  4.00%    … at this annual interest rate.
'Call exercised after    275  days
'Call exercised after: T=    0.75     years
'Loan repayment amount: L =  $18.53  That's $18.00, increased by 4.00% per year for 0.75 years.

Dim LOAN_REPAYMENT As Double 'Amount
Dim ANNUALIZED_RETURN As Double
Dim THEORETICAL_PREMIUM As Double

Dim TEMP_VECTOR As Variant

On Error GoTo ERROR_LABEL

THEORETICAL_PREMIUM = GENERALIZED_BLACK_SCHOLES_FUNC(SPOT, STRIKE, EXPIRATION, RATE, CARRY_COST, SIGMA, 1)
LOAN_REPAYMENT = Int(100 * AMOUNT_BORROWED * (1 + BORROWING_RATE) ^ CALL_EXERCISED_TENOR) / 100
ANNUALIZED_RETURN = ((STRIKE - LOAN_REPAYMENT) / (SPOT - AMOUNT_BORROWED - ACTUAL_PREMIUM)) ^ (1 / CALL_EXERCISED_TENOR) - 1

ReDim TEMP_VECTOR(1 To 7, 1 To 1)

TEMP_VECTOR(1, 1) = "You buy stock worth " & Format(SPOT, "#,0.00") & _
    " and borrow " & Format(AMOUNT_BORROWED, "#,0.00") & _
    ". Out-of-pocket cost to you is " & Format(SPOT, "#,0.00") & _
    "-" & Format(AMOUNT_BORROWED, "#,0.00") & "= " & _
    Format(SPOT - AMOUNT_BORROWED, "#,0.00") & ""

TEMP_VECTOR(2, 1) = "You sell a call for " & _
    Format(ACTUAL_PREMIUM, "#,0.00") & _
    ". Net Out-of-pocket cost is now " & _
    Format(SPOT - AMOUNT_BORROWED, "#,0.00") & _
    "-" & Format(ACTUAL_PREMIUM, "#,0.00") & " = " & _
    Format(SPOT - AMOUNT_BORROWED - ACTUAL_PREMIUM, "#,0.00") & "."

TEMP_VECTOR(3, 1) = "You 've borrowed " & _
    Format(AMOUNT_BORROWED, "#,0.00") & _
    " at " & Format(BORROWING_RATE, "0.00%") & " and, after " & _
    Format(CALL_EXERCISED_TENOR, "0.00") & " years the Call is exercised."

TEMP_VECTOR(4, 1) = "You sell the stock for " & _
    Format(STRIKE, "#,0.00") & _
    " and repay the " & Format(LOAN_REPAYMENT, "#,0.00") & _
    " loan. Net Dollars received is " & _
    Format(STRIKE - LOAN_REPAYMENT, "#,0.00") & "."

TEMP_VECTOR(5, 1) = "Your initial " & Format(SPOT - AMOUNT_BORROWED - _
    ACTUAL_PREMIUM, "#,0.00") & _
    " investment then returns " & Format(STRIKE - LOAN_REPAYMENT, "#,0.00") & _
    " after " & Format(CALL_EXERCISED_TENOR, "0.00") & " years."

TEMP_VECTOR(6, 1) = "Your Annualized Return is then [" & _
    Format(STRIKE - LOAN_REPAYMENT, "#,0.00") & _
    " / " & Format(SPOT - AMOUNT_BORROWED - ACTUAL_PREMIUM, "#,0.00") & _
    " ]^ (1/" & Format(CALL_EXERCISED_TENOR, "0.00") & _
    ") - 1 = " & Format(ANNUALIZED_RETURN, "0.00%")

TEMP_VECTOR(7, 1) = "Estimated Black-Scholes premium: " & _
    Format(THEORETICAL_PREMIUM, "#,0.00")

COVERED_CALL_RETURN_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
COVERED_CALL_RETURN_FUNC = Err.number
End Function


'Ratio Call Writing

'You write N Call Options and buy Sh shares of the stock.
'The Options have a Strike Price of K.
'You buy the stock at the current purchase price of S.
'The Option may be called at any stock price above K.
'This routine gives the Dollar and Percentage Gain (or Loss!)
'for each future stock price ST (assuming the option is called if ST > K).
'It also gives the WIDTH of the stock price interval where you'd make money.
'The WIDTH is expressed as a percentage of S, the purchase price of the stock
'and it could be plotted against the Strike Price ... also expressed as percentage of S.
'(Remember to show a big dot on the point corresponding to the data you entered)
'For example: Suppose the purchase price of the stock was S = $50,
'then WIDTH = 53% means 53% x $50 = $26.50 and K/S = 80% means K = $40.
'Although you may type in any Call Premium, Cp, you can also use the Black-Scholes premium BS.

Function COVERED_CALL_GAIN_LOSS_FUNC(ByVal PURCHASE_PRICE As Double, _
ByVal STRIKE_PRICE As Double, _
ByVal CALL_PREMIUM As Double, _
ByVal SELL_NO_CONTRACTS As Double, _
ByVal BUY_NO_SHARES As Long, _
Optional ByVal START_PRICE As Double = 10, _
Optional ByVal DELTA_PRICE As Double = 5, _
Optional ByVal NBINS As Long = 10, _
Optional ByVal OUTPUT As Integer = 2)

'Strike Price    K = $30.00
'Purchase Price  S = $40.00
'Sell how many Call contracts    N = 2
'Call Premium :  Cp =    $10.87
'Days to expiry      150
'Buy how many shares:    Sh =    100

Dim i As Long

Dim COST_VAL As Double 'Buy x shares, Sell y calls
Dim INCOME_VAL As Double 'income on calls

Dim TEMP_VAL As Double
Dim MIN_VAL As Double
Dim MAX_VAL As Double
Dim PRICE_VAL As Double

On Error GoTo ERROR_LABEL

COST_VAL = BUY_NO_SHARES * PURCHASE_PRICE - 100 * SELL_NO_CONTRACTS * CALL_PREMIUM
INCOME_VAL = 100 * SELL_NO_CONTRACTS * CALL_PREMIUM

If OUTPUT > 0 Then
    MIN_VAL = STRIKE_PRICE - (BUY_NO_SHARES * (STRIKE_PRICE - PURCHASE_PRICE) + 100 * SELL_NO_CONTRACTS * CALL_PREMIUM) / BUY_NO_SHARES
    MAX_VAL = STRIKE_PRICE - (BUY_NO_SHARES * (STRIKE_PRICE - PURCHASE_PRICE) + 100 * SELL_NO_CONTRACTS * CALL_PREMIUM) / (BUY_NO_SHARES - 100 * SELL_NO_CONTRACTS)
    If OUTPUT = 1 Then 'For stock prices in the range min to max you win
        COVERED_CALL_GAIN_LOSS_FUNC = Array(MIN_VAL, MAX_VAL)
        Exit Function
    Else '% Change from current stock price o
        COVERED_CALL_GAIN_LOSS_FUNC = Array(MIN_VAL / PURCHASE_PRICE - 1, MAX_VAL / PURCHASE_PRICE - 1)
        Exit Function
    End If
End If

ReDim TEMP_MATRIX(0 To NBINS, 1 To 3)
TEMP_MATRIX(0, 1) = "Stock Price"
TEMP_MATRIX(0, 2) = "Gain/Loss vs Stock price"
TEMP_MATRIX(0, 3) = "Gain/Loss Percentage"

PRICE_VAL = START_PRICE
For i = 1 To NBINS
    TEMP_MATRIX(i, 1) = PRICE_VAL

    If TEMP_MATRIX(i, 1) < STRIKE_PRICE Then
        TEMP_MATRIX(i, 2) = BUY_NO_SHARES * (TEMP_MATRIX(i, 1) - PURCHASE_PRICE) + 100 * SELL_NO_CONTRACTS * CALL_PREMIUM
    Else
        TEMP_VAL = 100 * SELL_NO_CONTRACTS - BUY_NO_SHARES
        If TEMP_VAL < 0 Then: TEMP_VAL = 0
        TEMP_MATRIX(i, 2) = 100 * SELL_NO_CONTRACTS * (STRIKE_PRICE + CALL_PREMIUM) - (PURCHASE_PRICE * BUY_NO_SHARES + TEMP_VAL * TEMP_MATRIX(i, 1))
    End If
    TEMP_MATRIX(i, 3) = IIf(COST_VAL > 0, TEMP_MATRIX(i, 2) / COST_VAL, 0)
    
    PRICE_VAL = PRICE_VAL + DELTA_PRICE
Next i

COVERED_CALL_GAIN_LOSS_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
COVERED_CALL_GAIN_LOSS_FUNC = Err.number
End Function

Function COVERED_CALL_WINNING_WIDTH_FUNC(ByVal PURCHASE_PRICE As Double, _
ByVal STRIKE_PRICE As Double, _
ByVal CALL_PREMIUM As Double, _
ByVal SELL_NO_CONTRACTS As Double, _
ByVal BUY_NO_SHARES As Long, _
ByVal VOLATILITY As Double, _
ByVal RISK_FREE_RATE As Double, _
ByVal EXPIRATION As Double, _
Optional ByVal CND_TYPE As Integer = 0, _
Optional ByVal START_PERCENT As Double = 0.25, _
Optional ByVal DELTA_PERCENT As Double = 0.1, _
Optional ByVal NBINS As Long = 10)

Dim i As Long

Dim KS_VAL As Double
Dim CURRENT_WIDTH As Double
Dim TEMP_PERCENT As Double

On Error GoTo ERROR_LABEL

KS_VAL = STRIKE_PRICE / PURCHASE_PRICE
CURRENT_WIDTH = (BUY_NO_SHARES * (STRIKE_PRICE - PURCHASE_PRICE) + 100 * SELL_NO_CONTRACTS * CALL_PREMIUM) * (1 / BUY_NO_SHARES - 1 / (BUY_NO_SHARES - 100 * SELL_NO_CONTRACTS)) / PURCHASE_PRICE

ReDim TEMP_MATRIX(0 To NBINS, 1 To 8)
TEMP_MATRIX(0, 1) = "Percent"
TEMP_MATRIX(0, 2) = "d1"
TEMP_MATRIX(0, 3) = "d2"
TEMP_MATRIX(0, 4) = "SS"
TEMP_MATRIX(0, 5) = "CCp"
TEMP_MATRIX(0, 6) = "Winning Width"
TEMP_MATRIX(0, 7) = "Current K/S"
TEMP_MATRIX(0, 8) = "Current Width"

TEMP_PERCENT = START_PERCENT
For i = 1 To NBINS
    TEMP_MATRIX(i, 1) = TEMP_PERCENT
    TEMP_MATRIX(i, 2) = (Log(1 / TEMP_MATRIX(i, 1)) + (RISK_FREE_RATE + _
                            VOLATILITY ^ 2 / 2) * EXPIRATION) / (VOLATILITY * Sqr(EXPIRATION))
    TEMP_MATRIX(i, 3) = TEMP_MATRIX(i, 2) - VOLATILITY * Sqr(EXPIRATION)
    TEMP_MATRIX(i, 4) = PURCHASE_PRICE
    TEMP_MATRIX(i, 5) = CND_FUNC(TEMP_MATRIX(i, 2), CND_TYPE) - _
                            TEMP_MATRIX(i, 1) * Exp(-RISK_FREE_RATE * EXPIRATION) * _
                            CND_FUNC(TEMP_MATRIX(i, 3), CND_TYPE)
    TEMP_MATRIX(i, 6) = (BUY_NO_SHARES * (TEMP_MATRIX(i, 1) - 1) + 100 * SELL_NO_CONTRACTS * _
                            TEMP_MATRIX(i, 5)) * (1 / BUY_NO_SHARES - 1 / _
                            (BUY_NO_SHARES - 100 * SELL_NO_CONTRACTS))
    TEMP_MATRIX(i, 7) = KS_VAL
    TEMP_MATRIX(i, 8) = CURRENT_WIDTH
    
    TEMP_PERCENT = TEMP_PERCENT + DELTA_PERCENT
Next i

COVERED_CALL_WINNING_WIDTH_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
COVERED_CALL_WINNING_WIDTH_FUNC = Err.number
End Function

Function COVERED_CALL_EXPECTED_PRICE_FUNC(ByVal PURCHASE_PRICE As Double, _
ByVal VOLATILITY As Double, _
ByVal ANNUAL_RETURN As Double, _
ByVal EXPIRATION As Double, _
Optional ByVal START_PRICE As Double = 10, _
Optional ByVal DELTA_PRICE As Double = 2.5, _
Optional ByVal NBINS As Long = 24, _
Optional ByVal OUTPUT As Integer = 1)

Dim i As Long

Dim TEMP1_SUM As Double
Dim TEMP2_SUM As Double
Dim TEMP_VAL As Double
Dim PRICE_VAL As Double

On Error GoTo ERROR_LABEL


ReDim TEMP_MATRIX(0 To NBINS, 1 To 3)
TEMP_MATRIX(0, 1) = "Stock Price"
TEMP_MATRIX(0, 2) = "Change X"
TEMP_MATRIX(0, 3) = "Probability"

TEMP1_SUM = 0
TEMP2_SUM = 0
PRICE_VAL = START_PRICE
For i = 1 To NBINS
    TEMP_MATRIX(i, 1) = PRICE_VAL
    TEMP_MATRIX(i, 2) = TEMP_MATRIX(i, 1) / PURCHASE_PRICE - 1
    TEMP_MATRIX(i, 3) = NORMDIST_FUNC(TEMP_MATRIX(i, 2), EXPIRATION * ANNUAL_RETURN, _
                        Sqr(EXPIRATION) * VOLATILITY, 0)
    
    PRICE_VAL = PRICE_VAL + DELTA_PRICE
    TEMP1_SUM = TEMP1_SUM + TEMP_MATRIX(i, 3)
    TEMP2_SUM = TEMP2_SUM + TEMP_MATRIX(i, 1) * TEMP_MATRIX(i, 3)
Next i

If OUTPUT = 0 Then
    COVERED_CALL_EXPECTED_PRICE_FUNC = TEMP_MATRIX
Else 'Expected Price
    PRICE_VAL = TEMP2_SUM / TEMP1_SUM
    TEMP_VAL = NORMDIST_FUNC(PRICE_VAL / PURCHASE_PRICE - 1, _
               EXPIRATION * ANNUAL_RETURN, Sqr(EXPIRATION) * VOLATILITY, 0)
    COVERED_CALL_EXPECTED_PRICE_FUNC = Array(PRICE_VAL, TEMP_VAL)
End If

Exit Function
ERROR_LABEL:
COVERED_CALL_EXPECTED_PRICE_FUNC = Err.number
End Function

'Assume Stock Price reaches >> $108.00 … and the Strike is $100.00
'You buy shares at $90.00 and sell calls at $10.01.
'If the stock goes to $108.00 you'd make $10.01 (for the option) + $100.00 for
'the stock. (That's the Strike Price.) Your profit (per share) would then be:
'($10.01 + $100.00 - $90.00) >>> $20.01   … less fees !!

Function COVERED_CALL_DELTA_FUNC(ByRef STOCK_PRICE_RNG As Variant, _
ByVal STRIKE_PRICE_RNG As Variant, _
ByVal EXPIRATION As Double, _
ByVal VOLATILITY As Double, _
ByVal RISK_FREE_RATE As Double, _
Optional ByVal DELTA_PRICE As Double = 1, _
Optional ByVal CND_TYPE As Integer = 0, _
Optional ByVal OUTPUT As Integer = 1)

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim BS1_PREMIUM As Double
Dim BS2_PREMIUM As Double

Dim TEMP_MATRIX As Variant
Dim STOCK_PRICE_VECTOR As Variant
Dim STRIKE_PRICE_VECTOR As Variant

On Error GoTo ERROR_LABEL

If IsArray(STOCK_PRICE_RNG) = True Then
    STOCK_PRICE_VECTOR = STOCK_PRICE_RNG
    If UBound(STOCK_PRICE_VECTOR, 1) = 1 Then
        STOCK_PRICE_VECTOR = MATRIX_TRANSPOSE_FUNC(STOCK_PRICE_VECTOR)
    End If
Else
    ReDim STOCK_PRICE_VECTOR(1 To 1, 1 To 1)
    STOCK_PRICE_VECTOR(1, 1) = STOCK_PRICE_RNG
End If
NROWS = UBound(STOCK_PRICE_VECTOR, 1)

If IsArray(STRIKE_PRICE_RNG) = True Then
    STRIKE_PRICE_VECTOR = STRIKE_PRICE_RNG
    If UBound(STRIKE_PRICE_VECTOR, 1) = 1 Then
        STRIKE_PRICE_VECTOR = MATRIX_TRANSPOSE_FUNC(STRIKE_PRICE_VECTOR)
    End If
Else
    ReDim STRIKE_PRICE_VECTOR(1 To 1, 1 To 1)
    STRIKE_PRICE_VECTOR(1, 1) = STRIKE_PRICE_RNG
End If
NCOLUMNS = UBound(STRIKE_PRICE_VECTOR, 1)

ReDim TEMP_MATRIX(1 To NROWS + 1, 1 To NCOLUMNS + 1)
For i = 1 To NROWS
    TEMP_MATRIX(1 + i, 1) = STOCK_PRICE_VECTOR(i, 1)
Next i
For j = 1 To NCOLUMNS
    TEMP_MATRIX(1, 1 + j) = STRIKE_PRICE_VECTOR(j, 1)
Next j

'------------------------------------------------------------------------------------------
Select Case OUTPUT
'------------------------------------------------------------------------------------------
Case 0
'------------------------------------------------------------------------------------------
    TEMP_MATRIX(1, 1) = "DELTA"
    For j = 1 To NCOLUMNS
        For i = 1 To NROWS
            BS1_PREMIUM = BLACK_SCHOLES_OPTION_FUNC(STOCK_PRICE_VECTOR(i, 1), _
                          STRIKE_PRICE_VECTOR(j, 1), EXPIRATION, _
                          RISK_FREE_RATE, VOLATILITY, 1, CND_TYPE)
            
            BS2_PREMIUM = BLACK_SCHOLES_OPTION_FUNC(STOCK_PRICE_VECTOR(i, 1) + _
                          DELTA_PRICE, STRIKE_PRICE_VECTOR(j, 1), _
                          EXPIRATION, RISK_FREE_RATE, VOLATILITY, 1, CND_TYPE)
        
            TEMP_MATRIX(i + 1, j + 1) = BS2_PREMIUM - BS1_PREMIUM
        Next i
    Next j
'------------------------------------------------------------------------------------------
Case Else
'------------------------------------------------------------------------------------------
    TEMP_MATRIX(1, 1) = "CALL / STOCK"
    For j = 1 To NCOLUMNS
        For i = 1 To NROWS
            BS1_PREMIUM = BLACK_SCHOLES_OPTION_FUNC(STOCK_PRICE_VECTOR(i, 1), _
                          STRIKE_PRICE_VECTOR(j, 1), EXPIRATION, _
                          RISK_FREE_RATE, VOLATILITY, 1, CND_TYPE)
            
            TEMP_MATRIX(i + 1, j + 1) = BS1_PREMIUM / STOCK_PRICE_VECTOR(i, 1)
        Next i
    Next j
'------------------------------------------------------------------------------------------
End Select
'------------------------------------------------------------------------------------------

COVERED_CALL_DELTA_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
COVERED_CALL_DELTA_FUNC = Err.number
End Function

'Covered Call Options
'There 's this Option strategy that I'd like to talk about. It goes like this ...
'1. You buy a bunch of shares of some stock, say IBM.
'it 'll cost you $90.78 a share.

'2. You then sell Call Options which expire in one month, with a Strike Price of $100.
'You 'll get $2.90 for the Option.

'3. If the Option is exercised within the next month, you'll get $100 for your stock (which
'you bought at $90.78). You got yourself $2.90 + $100 and you spent $90.78 for a profit of:
'$12.12 in one month.

'That 's a return of 12.12/90.78 ... about 13% in a month.

'4. If the Option is NOT exercised, you'd still have the stock and the $2.90 for having
'sold the Option. That 'll mean a gain of 2.90/90.78 ... about 3% in a month.

'And what about fees for selling Options? And what if the stock price drops like a rock?
'And what if ...? Yes, yes ... but wait until you hear the rest.
'Here 's a neat chart which shows the cost of a Call Option (as a percentage of the Stock Price):
'(It's an estimate obtained using the magic Black-Scholes formula.)
'it 'll depend upon the Strike Price (as a percentage of the Stock Price)
'... and the Volatility and some Risk-free rate.

'For the situation we described above, the Strike Price ($100) is about 110% of the Stock Price ($90.78)
'According to the chart, we'd expect an Option Price of about 1% of the Stock Price.

'Wait! In the IBM example the Option Price was ... uh ...
'It was about 3%.
'Yes, I know, it doesn't agree with Black-Scholes, but I did say that Black-Scholes
'was an "estimate".

'Anyway , i 'm not pushing Black-Scholes (which we may call BS, for short).
    
'Suppose we could find a stock where the Option Price (expiring in a month were) was 10% of
'the Stock Price (instead of a measly 1% or 3%).

'Are you still talking about Options with a Strike above the Stock Price?
'Yes. Out-of-the-money by 10% ... and expiring in a month.

'But then you'd make about 10% if the option is NOT exercised and about 20% if it is, right?
'Aha! You forgot fees.
'But, yes ... whether the option is exercised or not, your monthly gain should be significant.

'I thought you'd never consider out-of-the-money options.
'I assume you're talking about this. But we're now talking about selling, not buying.

'And what if your stock drops like a rock?
'Then you'll need this.

'Very funny ... but where do I find stocks with Call Options worth 10% of the Stock Price?
'Shhh ... that's a secret ...

'Okay, it goes like this:
'You type in today's Date, the month when the Call Option expires, the Strike price and the
'current Stock price.
'The spreadsheet tells you what the Option is worth ...according to Black-Scholes (BS).

'But don't you have to tell BS the Volatility of the stock?
'Uh ... yes, and BS also needs some Risk-free Rate.
'SO, you also type in the Actual value of the Option and then you ask the spreadsheet to vary the
'Volatility & Risk-free until its estimate is in good agreement.

'After the spreadsheet has identified appropriate Volatility and Risk-free Rate, you can now play.

'You can see how the ratio (Option price)/(Stock price) depends upon the ratio (Strike price)/(Stock price).

'You can see what you'd make if the Stock price reaches a certain value at expiry.
'it 'll be (Option price + Strike price - Stock price) if your Option gets called away.
'it 'll be just (Option price) if it doesn't ... and you get to keep the stock.

'Somebody has paid you for the privilege of buying your stock at the Strike price.
'If the stock actually exceeds that Strike, they'll buy your stock. If not, they won't.


'On that spreadsheet, it looks like the Call/Stock ratio is 11.5%?
'Yes, when Strike/Stock is 110%.

'That seems high.
'Yes, it pretty high. However, somebuddy told me about some other guys ... like these:
'    * SDS:
'          o Strike / Stock = 110
'          o Call/Stock = 9%
'    * DUG:
'          o Strike / Stock = 110
'          o Call/Stock = 12%
'    * DIG:
'          o Strike / Stock = 110
'          o Call/Stock = 14%


'>Uh ... what's that Delta stuff?
'The Delta value for an option gives you an indication of how much the option price would
'change for a $1.00 change in the stock price.

'If, for example, Delta = 0.25, then you might expect the option to increase by $0.25 if
'the stock price increased by $1.

'In the example illustrated by the spreadsheet, Delta = 0.48.
'So, if the stock increased from $90 to $91, one might expect the option to increase by $0.48,
'from $10 to $10.48.

'>And if the stock price increased by $10, then the option would increase by $4.80, from $10
'to $14.80, right? That 's a stretch.



'Reference:
'http://www.gummy-stuff.org/hedging.htm

Function COVERED_CALL_DELTA_HEDGING_FUNC(ByVal STOCK_PRICE As Double, _
ByVal STRIKE_PRICE As Double, _
ByVal VOLATILITY As Double, _
ByVal EXPECTED_RETURN As Double, _
ByVal RISK_FREE_RATE As Double, _
ByVal EXPIRATION As Double, _
Optional ByVal MIN_EXPIRATION As Double = 10 / 52, _
Optional ByVal DELTA_EXPIRATION As Double = 2 / 52, _
Optional ByVal NBINS_EXPIRATIONS As Long = 5, _
Optional ByVal MIN_STOCK_PRICE As Double = 50, _
Optional ByVal DELTA_STOCK_PRICE As Double = 1.2, _
Optional ByVal NBINS_STOCK_PRICES As Long = 26, _
Optional ByVal CND_TYPE As Integer = 0, _
Optional ByVal OUTPUT As Integer = 2)

'Stock Price   STOCK_PRICE =   $60.00
'Strike Price   STRIKE_PRICE =  $62.00
'Volatility VOLATILITY = 20
'Expected Return   Rtn = 10%
'Risk-free Rate   RISK_FREE_RATE =   5%
'Weeks to Maturity = 20

Dim i As Long
Dim j As Long
Dim TEMP_VAL As Double

Dim A_VAL As Double
Dim B_VAL As Double
Dim BS_VAL As Double
Dim DELTA_VAL As Double
Dim PRICE_VAL As Double
Dim TEMP_EXPIRATION As Double
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

ReDim TEMP_MATRIX(1 To NBINS_STOCK_PRICES + 1, 1 To NBINS_EXPIRATIONS + 1)


PRICE_VAL = MIN_STOCK_PRICE
For i = 1 To NBINS_STOCK_PRICES
    TEMP_MATRIX(i + 1, 1) = PRICE_VAL
    PRICE_VAL = PRICE_VAL + DELTA_STOCK_PRICE
Next i

TEMP_EXPIRATION = MIN_EXPIRATION
For j = 1 To NBINS_EXPIRATIONS
    TEMP_MATRIX(1, j + 1) = TEMP_EXPIRATION
    TEMP_EXPIRATION = TEMP_EXPIRATION + DELTA_EXPIRATION
Next j

BS_VAL = Int(100 * (STOCK_PRICE * CND_FUNC((Log(STOCK_PRICE / STRIKE_PRICE) + _
        (RISK_FREE_RATE + VOLATILITY ^ 2 / 2) * EXPIRATION) / (VOLATILITY * Sqr(EXPIRATION)), CND_TYPE) - _
        STRIKE_PRICE * Exp(-RISK_FREE_RATE * EXPIRATION) * _
        CND_FUNC((Log(STOCK_PRICE / STRIKE_PRICE) + (RISK_FREE_RATE + VOLATILITY ^ 2 / 2) * EXPIRATION) / _
        (VOLATILITY * Sqr(EXPIRATION)) - VOLATILITY * Sqr(EXPIRATION), CND_TYPE))) / 100
        
DELTA_VAL = CND_FUNC((Log(STOCK_PRICE / STRIKE_PRICE) + (RISK_FREE_RATE + _
            VOLATILITY ^ 2 / 2) * EXPIRATION) / (VOLATILITY * Sqr(EXPIRATION)), CND_TYPE)

Select Case OUTPUT
Case 0
    TEMP_MATRIX(1, 1) = "Distribution of Stock Prices"
    For i = 1 To NBINS_STOCK_PRICES
        For j = 1 To NBINS_EXPIRATIONS
            PRICE_VAL = TEMP_MATRIX(i + 1, 1)
            TEMP_EXPIRATION = TEMP_MATRIX(1, j + 1)
            A_VAL = (RISK_FREE_RATE - 0.5 * VOLATILITY ^ 2) / TEMP_EXPIRATION
            B_VAL = 1 / (2 * TEMP_EXPIRATION * VOLATILITY ^ 2)
            TEMP_VAL = B_VAL / PRICE_VAL * Exp(-B_VAL * (Log(PRICE_VAL / STOCK_PRICE) - A_VAL) ^ 2)
            TEMP_MATRIX(i + 1, j + 1) = TEMP_VAL
        Next j
    Next i
Case 1
    TEMP_MATRIX(1, 1) = "Delta Versus Stock Prices"
    For i = 1 To NBINS_STOCK_PRICES
        For j = 1 To NBINS_EXPIRATIONS
            PRICE_VAL = TEMP_MATRIX(i + 1, 1)
            TEMP_EXPIRATION = TEMP_MATRIX(1, j + 1)
            TEMP_VAL = CND_FUNC((Log(PRICE_VAL / STRIKE_PRICE) + (RISK_FREE_RATE + VOLATILITY ^ 2 / 2) * _
                        TEMP_EXPIRATION) / (VOLATILITY * Sqr(TEMP_EXPIRATION)), CND_TYPE)
            TEMP_MATRIX(i + 1, j + 1) = TEMP_VAL
        Next j
    Next i
Case Else
    TEMP_MATRIX(1, 1) = "Distribution of Gains"
    For i = 1 To NBINS_STOCK_PRICES
        For j = 1 To NBINS_EXPIRATIONS
            PRICE_VAL = TEMP_MATRIX(i + 1, 1)
            TEMP_EXPIRATION = TEMP_MATRIX(1, j + 1)
            
            If PRICE_VAL > STRIKE_PRICE + BS_VAL Then
                TEMP_VAL = DELTA_VAL * (PRICE_VAL - STOCK_PRICE) - _
                          (Int((PRICE_VAL * CND_FUNC((Log(PRICE_VAL / STRIKE_PRICE) + _
                          (RISK_FREE_RATE + VOLATILITY ^ 2 / 2) * TEMP_EXPIRATION) / _
                          (VOLATILITY * Sqr(TEMP_EXPIRATION)), CND_TYPE) - STRIKE_PRICE * _
                          Exp(-RISK_FREE_RATE * TEMP_EXPIRATION) * CND_FUNC((Log(PRICE_VAL / _
                          STRIKE_PRICE) + (RISK_FREE_RATE + VOLATILITY ^ 2 / 2) * _
                          TEMP_EXPIRATION) / (VOLATILITY * Sqr(TEMP_EXPIRATION)) - _
                          VOLATILITY * Sqr(TEMP_EXPIRATION), CND_TYPE))) - BS_VAL)
            Else
                TEMP_VAL = BS_VAL
            End If
            TEMP_MATRIX(i + 1, j + 1) = TEMP_VAL
        Next j
    Next i
End Select

COVERED_CALL_DELTA_HEDGING_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
COVERED_CALL_DELTA_HEDGING_FUNC = Err.number
End Function

'Suppose we invest in some stock or mutual fund or option or real estate or ...
'Unless the investment is risk-free, it may be good to invest in some other asset to
'counter the possibility that our first investment may go South.

'Suppose we invest in a stock currently worth $36.
'Suppose, further, that we estimate future stock prices/

'It Doesn 't matter. Just listen.
'Our estimate says there's a 30% probabillity that we'll lose money in N months.

'And 70% that we'll make money. Not bad, I'd say.
'What we want to do is invest in some other asset that is likely to go UP when our stock price
'goes down. We want an asset so that the distribution of our two-asset portfolio tends to
'counteract losses in the first asset and then reduce the standard deviation of our portfolio.

'Reduce the standard deviation without destroying your portfolio return, right?
'Well ... yes, tho' you may be willing to accept a somewhat smaller expected return in order
'to increase your probability of making money.
    
'Note that if we add an asset whose correlation with our stock is negative, the standard deviation
'will decrease. In general, most hedge funds attempt to reduce volatility (and risk) while attempting
'to deliver positive returns under all market conditions.

'For example:
'Suppose the stock had an Expected (or Mean) Annual Return of R and Standard Deviation of S, then a
'reasonable approximation for your compound annual growth rate is: CAGR = R - (1/2)S2 (See CAGR)

'Decrease S and you get a larger CAGR ... as we've noted above.

'Okay, now suppose we devote a fraction x of our portfolio to the stock and y = (1-x) to a second
'asset with Expected Return and Standard Deviation of A and B respectively.

'Our 2-asset portfolio would then have an Expected Return of x R + y A and (Standard Deviation)2 of
'x2S2 + y2B2+ 2 r x y S B
'Here, r is the correlation between the two assets.
'Okay, for our 2-asset portfolio our CAGR estimate would then be:
'CAGR = x R + y A - (1/2) { x2S2 + y2B2+ 2 r x y S B }

'Let's say the stock parameters are:
'R = 10%     S = 25%
'and our second asset has:
'A = 4%     B = 8%
'and the correlation varies from -100% to +100% (or -1.0 to +1.0).

'Note that, when the correlation is quite negative, we can reduce the volatility and increase the CAGR
'I like r = -1 Yes. Good luck in finding such a second asset ...
'However, if one could find such an asset and we kept 80% in our stock and put 20% into that second
'asset we'd reduce the volatility of our portfolio from 25% to about 18%.

'And you'd have a smaller expected return?
'Yes , It 'd be reduced from R = 10%   to   80%*R + 20%*A = 80%*10% + 20%*4% = 8.8%

'I prefer the larger expected return ... that R = 10%.
'The CAGR is a better representation of what you'd get. Did you notice that it just went up?
'Uh ... yes.
 
'Delta Hedging ... for Call Options

'suppose we 've written (and sold) a call option on a stock where the current stock parameters are:
'Stock Price S = $49
'Strike Price K = $50
'Stock Volatility V = 20%
'Expected Stock Return = 13%
'Weeks to Expiry T = 20 ... so Years to Expiry T = 20/52
'Risk-free Rate R = 5%

'The Black-Scholes option price for 100 shares (that's one contract) would be:
'Option contract Price = $240 ... meaning the call option is worth $2.40, but we're
'talking 100 shares, eh?
'Black-Scholes? Where'd that come from?
'It Doesn 't matter ... however the formula is:
'B-S Call Option price = S*NORMSDIST((LN(S/K)+(R+V^2/2)*T)/(V*SQRT(T))) -
'K*EXP(-R*T)*NORMSDIST((LN(S/K)+(R+V^2/2)*T)/(V*SQRT(T))-V*SQRT(T))

'The B-S formula doesn't involve the expected stock return!
'Interesting, eh?
'Anyway, suppose Sam buys our contract ... paying us $240.

'Sam will certainly NOT exercise the option immediately (asking us to sell him the stock at $50
'when it's available on the market for $49).

'Aah, but suppose the stock price goes to $52 and Sam exercises the option.
'we 'd sell Sam 100 shares at $50 and he could then sell those shares at $52. Sam's gain would be:
'[a]$5200 (from the sale of 100 shares @ $52) - $5000 (the cost of buying your 100 shares @ $50) -
'$240 (the cost of buying the contract) = - $40.

'Our gain would be:
'[b] $5000 (from the sale of 100 shares @ $50 to Sam) - $4900 (the cost of buying our 100 shares
'@ $49) + $240 (from the sale of the contract to Sam) = $340.

'Sam's crazy, right? Why would he exercise the option ... and lose money?
'He probably wouldn't.
'However, Sam can sell the contract at any time, so he keeps his eye on the stock price and notes
'how the value of the contract changes with the stock price. As the option price increases, the
'probability of Sam exercising his option increases.

'The following routines shows the option price for various stock prices, one week later
'(with 19 weeks to expiry)
'One week later?
'Yes. We use the B-S formula above, but with 19 weeks instead of 20.

'However, suppose we never really owned any shares ... so we wrote a "naked" option.
'If Sam exercises the option we must sell him 100 shares at $50.
'If the shares are worth $52 when Sam exercises the option, we've got to buy at $52 and sell to
'Sam at $50. Now our gain would be:
    
'[c] $5000 (from the sale of 100 shares @ $50) - $5200 (the cost of buying your 100 shares
'@ $52) + $240 (from the sale of the contract) = $40

'You make what Sam loses? Yes, for a naked option.
'Note that the terms in equation [a] are the negative of those in equation [c].
'Now comes the big problem: What if the price goes to, say $55 or $56 or ...
'You lose a bundle, eh?
'Yes, if we have to buy the stock at $55 or $56 ... and sell it at $50.
    
'Note that, if the stock price is S and the strike price is K and the option is worth C, our
'gain (or loss) is just: gain = K + C - S

'For each $1 increase in stock price S, our 100 shares give us a gain which decreases by $100.

'And where does the hedging come in?
'Good question. We want to hedge against the possibility that the stock price will go up.
'If we actually bought some additional stock, say N shares, then we could sell those at the
'higher price. Then we'd make some gains on the stock sale and that'd offset some of our option
'losses.

'How many additional shares would you buy?
'Aah , that 's where Delta Hedging comes in.
'First, recall that there's a number called Delta that measures how rapidly the option price will
'increase when the stock price goes up.

'Or decrease when the stock tanks, right?
'Uh ... yes, though we're mostly interested in having to buy stock at a higher price to cover our
'option commitment.

'So Here 's the scenario:
'1. We sell Sam a contract for $240 ... a 100-share contract at $2.40 option premium, as noted above.
'2. We calculate delta = 0.52 from the current parameters ... as given above.
'3. We borrow $2548 to buy 100*delta = 52 shares of the stock ... 52 shares at $49 per share = 2548.
'4. Suppiose that, in a week, the stock goes up to $52 and Sam asks for his 100 shares at $50.
'5. We run out and buy call options from Sally, at $4.16 ... the red dot in Figure 5.
'6. We exercise the contract we just purchased, get our 100 shares from Sally (worth $52) and pay Sam
'   his 100 shares.
'7. We sell our 52 shares of the stock at the current price of $52 ... and that's a gain of $3 per
'   share, hence a profit of 3*52 = $156.
'8. We repay the $2548 loan .

'So what 's our gain (or loss)?
'+$240   from the sale of the 100-share contract to Sam
'-$416   for the purchase of the 100-share contract from Sally
'+$156   from the sale of the 52 shares we bought at $49 and sold for $52
'= -$20  our net gain (or loss)

'How about the cost of borrowing?
'Yes.we 'll get to that in a moment.
'But you just lost $20!
'Yes.we 'll get to that too ... in a moment.

'Note that:
'* When the stock price increases and Sam exercises his option and asks for his shares, we buy an
'  option at the going rate from Sally.
'* We pay more to Sally than we got from Sam since the option increases by delta for each $1.00
'  increase in stock price.
'* However, we also buy delta shares of the stock ourself, so we make a profit when the stock goes up.
'* When the stock goes up by $X we lose delta*X on the option, but make delta*X on the stock.
'* Then ...

'But won't delta change? After all, it depends upon ...
'Yes, it depends upon volatility, time to expiry, etc. That means we change our stock position as
'the parameters change. that 's called ...

'Dynamic hedging, right?
'You got it ... and if you just sit there with your original stock purchase, doing nothing, it's
'static hedging. If we play our cards right we should be able to generate an almost risk-free
'strategy, like so:
'1. We sell an option for $C1 when the stock price is $S1.
'2. We borrow $D S1 to buy D shares at the current price of $S1 per share.
'Suppose the weekly interest rate is i   (where i = 0.00089 means 0.089%).
'The weekly interest is then   $ i D S1.
'3. N weeks pass and the stock price has increased to $S2 and somebuddy exercises the option we sold.
'4. We then buy an option at price $C2 ... on the same stock.
'5. We collect the shares from the option we bought to cover the option we sold.
'we 've just lost   $C2 - $C1 on the buying and selling of the options.
'6. We also sell the D shares we bought at step 2, making D ($S2 - $S1) on the sale.
'So far our gain is:   D ($S2 - $S1) - ($C2 - $C1).
'7. We pay the N-week interest on the loan in step 2, namely $N i D S1
'Our gain is now:   D ($S2 - $S1) - ($C2 - $C1) - $N i D S1.
'8. In order to make this strategy risk-free, we make this gain = $0; that is:
'D (S2 - S1) - (C2 - C1) - N i D S1 = 0
'9. Hence our purchase of shares in step 2 should be for D shares where:
'D = (C2 - C1) / (S2 - S1) + stuff

'Note that (C2 - C1) / (S2 - S1) is the (average) rate of change in option price as the stock prices changes.
'For small changes, that's just delta ... and that's the number of shares we should buy in step 2.
'Hence the name delta hedging.

'Uh ... yes, that stuff is:   N i D S1 / (S2 - S1).
'For our example above, it'd be about 0.75 shares.

'That'd cover the borrowing cost?
'Roughly, but remember that (C2 - C1) / (S2 - S1) is the average rate of change in option price when
'the stock changes from S1 to S2 ... shown in Figure 5B, for our example.

'we don 't know that average in advance ... so we buy delta shares.

'Delta = NormSDist((Ln(S / K) + (r + V ^ 2 / 2) * T) / (V * SQRT(T)))

'And that delta hedging is risk-free? Well ... close.
'So, you think you understand this deta hedging stuff? Uh ... barely.
    
'Here 's the scheme we're investigating:

'suppose we 've written (and sold) a call option on a stock where the current stock parameters are:

'stock Price = S
'Strike Price = K
'stock VOLATILITY = V
'Expected Stock Return = R
'Years to Expiry = T
'Risk-free Rate = Rf

'We use, as the price of the option, the Black-Scholes formula:
'[A] Option price C = S*NORMSDIST((LN(S/K)+(Rf+V^2/2)*T)/(V*SQRT(T))) -
'    K*EXP(-R*T)*NORMSDIST((LN(S/K)+(Rf+V^2/2)*T)/(V*SQRT(T))-V*SQRT(T))

'We calculate Delta, the rate of change of option price with respect the stock price , via:
'[B] Delta = dC / dS = NormSDist((Ln(S / K) + (Rf + V ^ 2 / 2) * T) / (V * SQRT(T)))
'1. We sell an option for $C when the stock price is $S.
'2. We buy D shares at the current price of $S per share, borrowing $D S to do so.
'3. The weekly interest rate for borrowing is i so the weekly cost of borrowing is i D S
'   (where i = 0.0008 means 0.08%).
'4. N weeks pass and the stock price has increased by $?S ... and the option is exercised.
'5. We then buy an option at a price which is larger than $C by an amount $?C ... on the same stock.
'6. We collect the shares from the option we bought to cover the option we sold.
'we 've just lost   $?C on the buying and selling of the options.
'7. But then we sell the D shares we bought at step 2, making D ?S on the sale.
'8. We pay the N-week interest on the loan in step 2, namely $N i D S
'Our net gain is:   D ?S - ?C - N i D S
'9. So as not to lose money, we choose D to make this net gain = $0; that is:
'D ?S - ?C - N i D S = 0
'10. Hence our purchase of shares in step 2 should be for D shares where:
'D = ?C / (?S - N i S )

'If, for example, N = 10 weeks and i = 0.0008 (corresponding to about 4% annual interest
'rate on our borrowing), then N i S = 0.008S or 0.8% of the stock price S.
'we 'll assume this is much smaller than the change in stock price ?S ... so we'll ignore it.

'Hence the number of shares we should hold (in step 2, above) is: D = ?C / ?S.

'And that's delta, for the stock, right?
'Yes, for small changes in stock price, since delta is actually dC / dS.

'C / ?S or dC / dS. What's the difference?
'In a car trip, ?C / ?S is like the average speed of your car and dC / dS is the instantaneous
'speed ... as noted on the speedometer.

'Anyway, since we are to hold sufficient stock to cover the loss in option trading, we should buy
'and sell stock as the weeks go by... since delta changes.

'And you're ignoring the cost of borrowing?
'Yes, to make things simple ... just to get a flavour of this delta hedging stuff.

'Okay, suppose our parameters are:
'Stock Price = S = $60
'Strike Price = K = $62
'stock VOLATILITY = V = 20
'Years to Expiry = T = 20/52 (meaning 20 weeks)
'Risk-free Rate = Rf = 4%
'and we sell 10 contracts (worth1000 shares of stock).

'Initially:
'C = $2.59
'Delta = 0.481
'and we'd buy (initiallly) 1000*0.481 = 481 shares of stock at $60 ... borrowing the money to do so.
'As the weeks go by, delta changes.
'The stock price S may change and the time to maturity T certainly changes ... see magic formula [B], above.
'the number of shares we should own with 15, 10 then 5 weeks to go before the option expires. It's a
'chart of 1000*delta.

'But you don't know the future stock price!
'True, but is interesting nevertheless. We just buy or sell our stock holdings to match the chart ...
'depending upon what the stock price happens to be at the time.
'For example:
'At 15 weeks to go and a stock price is $65 we should own about 730 shares.
'At 10 weeks to go and a stock price is $67 we should own about 840 shares.
'At 5 weeks to go and a stock price is $64 we should own about 735 shares.

'Yeah, I see the magenta dots.
    
'Now, if we could generate a distribution of stock prices at some time in the future, we could
'generate a distribution for delta ...

'Well ... so far I have something like this:
'You don 't use the expected return, do you?
'No. I just stuck it in the parameters as decoration.
'Where'd you get that stock price distribution?
'I stole it from here.
'That delta distribution looks like a cumulative probability distribution.
'Doesn 't it? Of course, in magic formula [B] above, that NORMSDIST function is a cumulative distribution.
'What if I make a mistake in estimating V, the volatility?
'don 't make a mistake else your delta-chart will change ... like this
'And if you buy the option, what then?
'You can short the stock ... delta-shares worth, to hedge against the stock dropping in price.
'Indeed, it seems to make more sense if you're the guy who bought the option and are worried about
'its value dropping as the stock price drops.

'And that's the way delta hedging works?
'Uh ... how would I know? I just regurgitate what I find on the Net.
'If the objective is to lose no money then I have another, simpler method
'Very funny. However, 99% of delta hedging is done by market makers (Norbert tells me). Their gains
'are elsewhere. They just don't want to lose anything when moving options between buyers and sellers.

'What's that funny chart ... Distribution of Gains after 2 weeks?
'it 's assumed that nobuddy exercises the option you write until the stock price exceeds K + C.
'In the example illustrated, that's $64.59
'If the stock price is less than this, you just make the money you got by selling the option, namely $2.59.


