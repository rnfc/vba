Attribute VB_Name = "FINAN_ASSET_SYSTEM_MOMENT_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'Momentum Investing ... maybe
'So there 's this strategy called Reverse Scale which is (sorta) the reverse of the usual
'strategy and is more like "momentum" investing and ...

'You know: Buy when the stock is down and Sell when it's up.
'That 's buy low, sell high, right?
'Yes, but this Reverse Scale Strategy goes something like this:
'1. Start with some investment, say $A, when the stock price is, say, Po.
'2. Pick some increase, say 30%. Then, whenever the price increases by 30%, buy another $A worth of stock.
'3. Continue repeating step #2 every time the stock increases by 30%.
'4. Note that there are decision points, namely when these prices occur: (1.3)Po, (1.3)2Po, (1.3)3Po, etc. etc.
'5. You Buy $A worth of stock whenever the stock hits such a price.
'6. However, if the stock retreats to a previous decision point, Sell everything.

'Well, the strategy suggests that, after selling everything, you start with a different stock,
'repeating the above steps.

'How do you find that "different" stock? Throw darts?
'Not at all. There are criteria for selecting stocks.
'however , I 'd like to talk about a modified strategy.

'We stick with the same stock and Sell if the stock retreats by 30% from the last decision price.
'Then, after selling everything, we Buy again when the price rises to the next higher decision point.

'Consider weekly closing prices for GE stock, from April, 1996 to April, 2006.
'That 's a bit over 500 weeks.
'We start by buying $500 dollars worth of stock at $10.69 ... in April, 1996.
'We pick some increase, say 30% and identify the magic price targets:
'10.69*(1.3) = 13.90, 13.90*(1.3) = 18.07, 18.07*(1.3) = 23.49
'and 30.53, 39.69, 51.60 etc.
'These are shown in red in Figure 1.
'We add another $500 whenever we exceed one of these.

'Notice, however, that the price dropped below the last magic Buy price of $39.69 at week 249, so
'we sold everything.
    
'You can start with a Portofolio which is different than the amount of stock you Buy at each
'Decision Price. In the picture, they're both $500.

'Just in case you want to Sell after a Decrease that's different than one of the prices at which you Buy.
'If you make them the same, like 30%, then you get a single set of Decision Prices.
'If they're different, the Buy and Sell decision prices are different.

'http://www.gummy-stuff.org/momentum.htm
'http://www.investopedia.com/university/fiveminute/fiveminute7.asp

Function ASSET_MOMENTUM_SIGNAL_FUNC(ByVal TICKER_STR As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal END_DATE As Date, _
Optional ByVal BUY_PERCENT As Double = 0.03, _
Optional ByVal SELL_PERCENT As Double = 0.02, _
Optional ByVal INITIAL_SYSTEM As Double = 1000, _
Optional ByVal BUY_CASH As Double = 1000, _
Optional ByVal OUTPUT As Integer = 0)

'INITIAL_SYSTEM: Portfolio $1,000 initially
'BUY_CASH: Buy $1,000 shares

Dim i As Long
Dim j As Long
Dim k As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

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
NCOLUMNS = UBound(DATA_MATRIX, 2)

'-----------------------------------------------------------------------------
ReDim TEMP_MATRIX(0 To NROWS, 1 To 15)
'-----------------------------------------------------------------------------
TEMP_MATRIX(0, 1) = "DATE"
TEMP_MATRIX(0, 2) = "OPEN"
TEMP_MATRIX(0, 3) = "HIGH"
TEMP_MATRIX(0, 4) = "LOW"
TEMP_MATRIX(0, 5) = "CLOSE"
TEMP_MATRIX(0, 6) = "VOLUME"
TEMP_MATRIX(0, 7) = "ADJ CLOSE"

'Decision Prices Action  Trade Price Buy Sell    Shares  Portfolio   # BUYs

TEMP_MATRIX(0, 8) = "DECISION PRICE"
TEMP_MATRIX(0, 9) = "ACTION"
TEMP_MATRIX(0, 10) = "TRADE PRICE"
TEMP_MATRIX(0, 11) = "BUY PLOT"
TEMP_MATRIX(0, 12) = "SELL PLOT"
TEMP_MATRIX(0, 13) = "NO SHARES"
TEMP_MATRIX(0, 14) = "SYSTEM BALANCE"

i = 1
For j = 1 To NCOLUMNS: TEMP_MATRIX(i, j) = DATA_MATRIX(i, j): Next j
TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 6) / 10000
TEMP_MATRIX(i, 8) = TEMP_MATRIX(i, 7)
TEMP_MATRIX(i, 9) = ""
TEMP_MATRIX(i, 10) = TEMP_MATRIX(i, 7) 'Initial Price
TEMP_MATRIX(i, 11) = ""
TEMP_MATRIX(i, 12) = ""
TEMP_MATRIX(i, 13) = Int(INITIAL_SYSTEM / TEMP_MATRIX(i, 10))
TEMP_MATRIX(i, 14) = INITIAL_SYSTEM
TEMP_MATRIX(i, 15) = ""

k = 0

For i = 2 To NROWS
    For j = 1 To NCOLUMNS: TEMP_MATRIX(i, j) = DATA_MATRIX(i, j): Next j
    TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 6) / 10000
    TEMP_MATRIX(i, 8) = ""
    
    If TEMP_MATRIX(i, 7) / TEMP_MATRIX(i - 1, 8) > 1 + BUY_PERCENT Then
        TEMP_MATRIX(i, 9) = "BUY"
    Else
        If (TEMP_MATRIX(i, 7) / TEMP_MATRIX(i - 1, 8) < 1 / (1 + SELL_PERCENT) And _
            TEMP_MATRIX(i - 1, 13) > 0) Then
            TEMP_MATRIX(i, 9) = "SELL"
        Else
            TEMP_MATRIX(i, 9) = ""
        End If
    End If
    
    If TEMP_MATRIX(i, 7) > TEMP_MATRIX(i - 1, 8) * (1 + BUY_PERCENT) Then
        TEMP_MATRIX(i, 8) = TEMP_MATRIX(i - 1, 8) * (1 + BUY_PERCENT)
    Else
        If (TEMP_MATRIX(i, 9) = "SELL" And TEMP_MATRIX(i, 7) < TEMP_MATRIX(i - 1, 8) / _
           (1 + SELL_PERCENT)) Then
            TEMP_MATRIX(i, 8) = TEMP_MATRIX(i - 1, 8) / (1 + SELL_PERCENT)
        Else
            TEMP_MATRIX(i, 8) = TEMP_MATRIX(i - 1, 8)
        End If
    End If

    '-------------------------------------------------------------------------------
    If (TEMP_MATRIX(i, 9) = "BUY" Or TEMP_MATRIX(i, 9) = "SELL") Then
        TEMP_MATRIX(i, 10) = TEMP_MATRIX(i, 7)
    Else
        TEMP_MATRIX(i, 10) = TEMP_MATRIX(i - 1, 10)
    End If
    
    '-------------------------------------------------------------------------------
    If TEMP_MATRIX(i, 9) = "BUY" Then 'Buy Plot
        TEMP_MATRIX(i, 11) = TEMP_MATRIX(i, 10)
    Else
        TEMP_MATRIX(i, 11) = "" '-10
    End If
    
    If TEMP_MATRIX(i, 9) = "SELL" And TEMP_MATRIX(i - 1, 13) > 0 Then 'Sell Plot
        TEMP_MATRIX(i, 12) = TEMP_MATRIX(i, 10)
    Else
        TEMP_MATRIX(i, 12) = "" '-10
    End If
    '-------------------------------------------------------------------------------
    
    If (TEMP_MATRIX(i - 1, 13) > 0 And TEMP_MATRIX(i, 9) = "BUY") Then
        TEMP_MATRIX(i, 13) = Int(TEMP_MATRIX(i - 1, 13) + BUY_CASH / TEMP_MATRIX(i, 10))
    Else
        If (TEMP_MATRIX(i - 1, 13) = 0 And TEMP_MATRIX(i, 9) = "BUY") Then
            TEMP_MATRIX(i, 13) = Int(TEMP_MATRIX(i - 1, 14) / TEMP_MATRIX(i, 10))
        Else
            If TEMP_MATRIX(i, 9) = "SELL" Then
                TEMP_MATRIX(i, 13) = 0
            Else
                TEMP_MATRIX(i, 13) = Int(TEMP_MATRIX(i - 1, 13))
            End If
        End If
    End If
    
    If TEMP_MATRIX(i, 13) > 0 Then
        TEMP_MATRIX(i, 14) = TEMP_MATRIX(i, 13) * TEMP_MATRIX(i, 7)
    Else
        If (TEMP_MATRIX(i, 9) = "SELL" And TEMP_MATRIX(i, 13) = 0) Then
            TEMP_MATRIX(i, 14) = TEMP_MATRIX(i - 1, 13) * TEMP_MATRIX(i, 7)
        Else
            TEMP_MATRIX(i, 14) = TEMP_MATRIX(i - 1, 14)
        End If
    End If
    
    If TEMP_MATRIX(i, 9) = "BUY" Then
        TEMP_MATRIX(i, 15) = 1
        k = k + 1
    Else
        TEMP_MATRIX(i, 15) = 0
    End If
Next i

TEMP_MATRIX(0, 15) = "NO BUYS = " & k

Select Case OUTPUT
Case 0
    ASSET_MOMENTUM_SIGNAL_FUNC = TEMP_MATRIX
Case Else 'Growth
    ASSET_MOMENTUM_SIGNAL_FUNC = (TEMP_MATRIX(NROWS, 14) / ((BUY_CASH * k) + INITIAL_SYSTEM)) - 1
End Select

Exit Function
ERROR_LABEL:
ASSET_MOMENTUM_SIGNAL_FUNC = Err.number
End Function
