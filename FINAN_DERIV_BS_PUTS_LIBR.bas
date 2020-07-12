Attribute VB_Name = "FINAN_DERIV_BS_PUTS_LIBR"
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


Function PUT_OPTIONS_BREAK_POINTS_FUNC(ByRef STRIKE_PRICES_RNG As Variant, _
ByRef PREMIUMS_RNG As Variant, _
ByRef CONTRACTS_RNG As Variant, _
ByRef SELL_BUY_DUMMY_RNG As Variant, _
Optional ByVal MIN_PRICE As Double = 60, _
Optional ByVal MAX_PRICE As Double = 82, _
Optional ByVal DELTA_PRICE As Double = 0.5, _
Optional ByVal OUTPUT As Integer = 1)

Dim i As Long
Dim j As Long
Dim NROWS As Long
Dim NBINS As Long

Dim TEMP_VAL As Double
Dim PRICE_VAL As Double
Dim BREAK_EVEN_PRICE As Double
Dim STRIKE_PRICES_VECTOR As Variant
Dim PREMIUMS_VECTOR As Variant
Dim CONTRACTS_VECTOR As Variant
Dim SELL_BUY_DUMMY_VECTOR As Variant

Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

STRIKE_PRICES_VECTOR = STRIKE_PRICES_RNG
If UBound(STRIKE_PRICES_VECTOR, 1) = 1 Then
    STRIKE_PRICES_VECTOR = MATRIX_TRANSPOSE_FUNC(STRIKE_PRICES_VECTOR)
End If
NROWS = UBound(STRIKE_PRICES_VECTOR, 1)

PREMIUMS_VECTOR = PREMIUMS_RNG 'market prices or black scholes premium
If UBound(PREMIUMS_VECTOR, 1) = 1 Then
    PREMIUMS_VECTOR = MATRIX_TRANSPOSE_FUNC(PREMIUMS_VECTOR)
End If
If NROWS <> UBound(PREMIUMS_VECTOR, 1) Then: GoTo ERROR_LABEL

CONTRACTS_VECTOR = CONTRACTS_RNG 'number of contracts
If UBound(CONTRACTS_VECTOR, 1) = 1 Then
    CONTRACTS_VECTOR = MATRIX_TRANSPOSE_FUNC(CONTRACTS_VECTOR)
End If
If NROWS <> UBound(CONTRACTS_VECTOR, 1) Then: GoTo ERROR_LABEL

SELL_BUY_DUMMY_VECTOR = SELL_BUY_DUMMY_RNG '+1 Buy / -1 Sell
If UBound(SELL_BUY_DUMMY_VECTOR, 1) = 1 Then
    SELL_BUY_DUMMY_VECTOR = MATRIX_TRANSPOSE_FUNC(SELL_BUY_DUMMY_VECTOR)
End If
If NROWS <> UBound(SELL_BUY_DUMMY_VECTOR, 1) Then: GoTo ERROR_LABEL

NBINS = Int((MAX_PRICE - MIN_PRICE) / DELTA_PRICE) + 1

ReDim TEMP_MATRIX(0 To NBINS, 1 To 3)
TEMP_MATRIX(0, 1) = "STOCK_PRICE"
TEMP_MATRIX(0, 2) = "PROFIT"
TEMP_MATRIX(0, 3) = "BREAK-EVEN"

BREAK_EVEN_PRICE = 0
PRICE_VAL = MIN_PRICE
For i = 1 To NBINS
    TEMP_MATRIX(i, 1) = PRICE_VAL
    TEMP_MATRIX(i, 2) = 0
    For j = 1 To NROWS
        TEMP_VAL = TEMP_MATRIX(i, 1) - STRIKE_PRICES_VECTOR(j, 1)
        TEMP_MATRIX(i, 2) = TEMP_MATRIX(i, 2) + (CONTRACTS_VECTOR(j, 1) * SELL_BUY_DUMMY_VECTOR(j, 1) * ((IIf(TEMP_VAL > 0, 0, TEMP_VAL) + PREMIUMS_VECTOR(j, 1))))
    Next j
    TEMP_MATRIX(i, 2) = TEMP_MATRIX(i, 2) * 100
    
    If i > 1 Then
        If TEMP_MATRIX(i - 1, 2) * TEMP_MATRIX(i, 2) <= 0 Then
            TEMP_MATRIX(i, 3) = TEMP_MATRIX(i - 1, 1) - TEMP_MATRIX(i - 1, 2) * (TEMP_MATRIX(i, 1) - TEMP_MATRIX(i - 1, 1)) / (TEMP_MATRIX(i, 2) - TEMP_MATRIX(i - 1, 2))
            BREAK_EVEN_PRICE = TEMP_MATRIX(i, 3)
            If OUTPUT = 1 Then: Exit For
        Else
            TEMP_MATRIX(i, 3) = ""
        End If
    Else
        TEMP_MATRIX(i, 3) = ""
    End If
    PRICE_VAL = PRICE_VAL + DELTA_PRICE
Next i

Select Case OUTPUT
Case 0
    PUT_OPTIONS_BREAK_POINTS_FUNC = TEMP_MATRIX
Case 1
    If BREAK_EVEN_PRICE <> 0 Then
        PUT_OPTIONS_BREAK_POINTS_FUNC = BREAK_EVEN_PRICE
    Else
        PUT_OPTIONS_BREAK_POINTS_FUNC = "Change Starting $X and/or dX"
    End If
Case Else
    PUT_OPTIONS_BREAK_POINTS_FUNC = Array(TEMP_MATRIX, BREAK_EVEN_PRICE)
End Select

Exit Function
ERROR_LABEL:
PUT_OPTIONS_BREAK_POINTS_FUNC = Err.number
End Function

'You Sell and/or Buy up to XYZ PUT options (using either bs = 1 or bs = -1)
'with Strike Prices SP1, SP2, etc. Premiums Pr1, Pr2, etc.
'and Number of Contracts Nr1, Nr2, etc.
'What 's your Break Even Stock Price ?


'Buying and Selling PUT options (up to x contracts per stock) ... and what the
'formula was for the Break Even stock price.

'First we note the following, for the Buyer of a PUT Option:

'If the Strike Price of a PUT option is SP, then a person buying a PUT will be able to sell it for $SP
'(anytime she wishes, before it expires).
'For this privilege, she pays a premium of Pr.
'If the Stock Price, X, is less than SP and the buyer exercises her option, she sells the stock
'at SP and can immediately buy it at X.
'Her profit is then (SP - X) - Pr   subtracting the cost of buying the option: Pr
'If the Stock Price is never less than the Strike Price, then she won't sell the stock at
'the X (since it's worth more, namely SP, in the market)
'In this case, her profit is 0 - Pr   which, of course, is a loss!
'In general, then, her PROFIT is: MAX(SP - X, 0) - Pr   where MAX means the MAXimum of the two values


'Now, for the Seller ...

'You mean the guy who writes the option?
'Yes ... but I'll call him the Seller. As I was saying:

'Now, for the Seller of a PUT Option:

'The seller of the PUT receives the premium of Pr   regardless of whether the option is exercised.
'If the Stock Price, X, is less than SP and the option buyer exercises her option, the seller of
'the PUT must buy the stock at SP even though it's only worth X in the market.
'His profit is then (X - SP) + Pr   adding the loss (X - SP) from the Premium he receives
'If the Stock Price is never less than the Strike Price, then he won't be buying any stock (!)
'In this case, his profit is just the premium, namely: Pr.
'In general, then, his PROFIT is: MIN(X - SP, 0) + Pr   where MIN means the Minimum of the two values

'Note that MAX(SP - X, 0) = - MIN(X - SP, 0) so that, for the Buyer, the PROFIT can also be written
'as:   - MIN(X - SP, 0) - Pr
'Comparing the two profit formulas, we can write the profit for both the Buyer and the Seller as:
'PROFIT= bs {MIN(X - SP, 0) + Pr}

'where bs = -1 for the Buyer
'and bs = 1 for the Seller

'Yeah ... for buysell.
'Note that the Seller's profit is the Buyer's loss ... and vice versa.
'Note, too, that we could also write:   PROFIT = bs {- MAX(SP - X, 0) + Pr}.

'What about fees and commissions and ...? we 're ignoring them
'So, what's that Break Even Stock Price?
'that ain 't easy. If you Buy and Sell several options, the PROFIT chart is a series of
'connected line segments.
'For Break Even, we want the Stock Price, X, where PROFIT = 0.

'Suppose we 're Selling (or "writing") four put options (so bs = 1), each involving Nr = 10
'contracts (a "contract" is 100 stock shares) where:

'The Strike Prices are SP1 = $65, SP2 = $70, SP3 = $75 and SP4 = $80
'The Premiums are Pr1 = $0.25, Pr2 = $1.10, Pr3 = $3.50 and Pr4 = $8.00
'The Number of contracts are Nr1 = 10, Nr2 = 10, Nr3 = 10 and Nr4 = 10

'The chart of PROFIT is the sum of a bunch of terms, like:
'100*bs Nr {MIN(X - SP, 0) + Pr}     with bs = 1 and Nr the Number of contracts.
'In other words, for our four options, it's:
'100* {Nr1[ MIN(X-SP1, 0)+Pr1 ] + Nr2[ MIN(X-SP2, 0)+Pr2 ]+ Nr3[ MIN(X-SP3, 0)+Pr3 ] +
'Nr4[ MIN(X-SP4, 0)+Pr4 ] }

'Note the strange, broken-line chart of PROFIT  - the breaks occur at the Strike Prices
'Each segment is a straight line so that we may find the place where PROFIT = 0 by identifying
'where the PROFIT changes sign.

'Suppose that occurs at (X,0), between (X1, Y1) and (X2, Y2).
'Then (because the graph is a straight line segment between Strike Prices), so we need to
'equates slopes, getting
'(0 - Y1) / (X - X1) = (Y2 - Y1) / (X2 - X1)
'so the Break Even is (solving for X):
'X = X1 - Y1 * (X2 - X1) / (Y2 - Y1)

'Can there be more than one break even?
'Uh ... yes, but if you're careful you can get 'em all

