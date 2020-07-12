Attribute VB_Name = "FINAN_ASSET_SYSTEM_DAYS_BACK"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'a Buy & Sell Strategy
'i 've always felt that people over-react to good or bad news ... especially stock traders.
'Earnings come in a penny below estimates, everybody yells "the sky is falling!" and the stock
'plummets ... only to quickly recover.

'So I figured I could take advantage of that as follows:

'I look for a high-volatility stock. One that often changes by 3 or 4% in a day.
'If the closing price is significantly lower than the open, I buy at the close ... figuring that
'it's an over-reaction and the stock will recover quickly.
'If the close is significantly higher than the open, I sell at the close.

'You always buy or sell at the close?
'Yes. I have my finger on the Buy / Sell button when the market closes.

Function ASSET_DAYS_BACK_SIGNAL_FUNC(ByVal TICKER_STR As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal END_DATE As Date, _
Optional ByVal BUY_PERCENT As Double = -0.02, _
Optional ByVal SELL_PERCENT As Double = 0.02, _
Optional ByVal INITIAL_SHARES As Double = 1000, _
Optional ByVal SHARES_TRADE As Double = 200, _
Optional ByVal INITIAL_CASH As Double = 5000, _
Optional ByVal DAYS_BACK As Long = 1, _
Optional ByVal OUTPUT As Integer = 1)

'Initial Shares: We start with a certain number of Shares

'Buy Percent: We buy whenever the Close is smaller than
'a previous Open by a certain percentage

'Sell Percent: We sell whenever the Close is larger
'than a previous Open by a certain percentage

'Shares Trade: buy or sell a fixed number of Shares (at
'the closing price)

'Initial Cash: start with a bunch of Cash.
'Of course, we don't buy if we don't have enough Cash to
'handle the trade and we don't sell shares unless we have
'that many shares to sell.

'Days Back: We might look for dramatic changes over 1, 2 or 3 days, comparing today's
'Close with the Open a few days ago. In this example, we're comparing today's
'Close with yesterday's Open (back 1 day) in order to determine whether we
'should BUY or SELL at today's Close.

'Reference: http://www.gummy-stuff.org/buy-sell.htm

Dim i As Long
Dim j As Long
Dim k As Long
Dim NROWS As Long

Dim MEAN_VAL As Double
Dim VOLAT_VAL As Double

Dim Z_VAL As Double

Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL


'We start with a certain number of Shares (Example: 1000, in the picture above)
'We also start with a bunch of Cash (Example: $5,000)
'We always buy or sell a fixed number of Shares (at the closing price) (Example: 200)
'We BUY whenever the Close is smaller than a previous Open by a certain percentage (Example: -2%)
'We SELL whenever the Close is larger than a previous Open by a certain percentage (Example: 2%)
'Of course, we don't buy if we don't have enough Cash to handle the trade and we don't sell 200
'shares unless we have that many shares to sell.

'Compared to a previous Open?
'Yes, we might look for dramatic changes over 1, 2 or 3 days, comparing today's Close with the Open
'a few days ago.

'In the example, we're comparing today's Close with yesterday's Open (back 1 day) in order to determine
'whether we should BUY or SELL at today's Close.

'Notice, too, that we've sold all our shares by Mar/06 then start buying as MSFT decreases in May & Jun/06.
'Then we're selling again as MSFT recovers and end up selling all our shares by Dec/06 (which accounts
'for the "flat" portfolio graph).

If IsArray(TICKER_STR) = False Then
    DATA_MATRIX = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, _
                  START_DATE, END_DATE, "DAILY", "DOHLCVA", True, _
                  False, True)
Else
    DATA_MATRIX = TICKER_STR
End If
NROWS = UBound(DATA_MATRIX, 1)

ReDim TEMP_MATRIX(0 To NROWS, 1 To 16)

TEMP_MATRIX(0, 1) = "Date"
TEMP_MATRIX(0, 2) = "Open"
TEMP_MATRIX(0, 3) = "High"
TEMP_MATRIX(0, 4) = "Low"
TEMP_MATRIX(0, 5) = "Close"
TEMP_MATRIX(0, 6) = "Volume"
TEMP_MATRIX(0, 7) = "Adj. Close"

TEMP_MATRIX(0, 8) = "Returns"

TEMP_MATRIX(0, 9) = "Volume/1000"
TEMP_MATRIX(0, 10) = "Open"
TEMP_MATRIX(0, 11) = "Close"
TEMP_MATRIX(0, 12) = "%Change"
'---------------------------------------------
TEMP_MATRIX(0, 13) = "Shares"
TEMP_MATRIX(0, 14) = "Cash"
TEMP_MATRIX(0, 15) = "System"
TEMP_MATRIX(0, 16) = "Returns"
'---------------------------------------------

For i = 1 To NROWS
    For j = 1 To 7: TEMP_MATRIX(i, j) = DATA_MATRIX(i, j): Next j
    If i <> 1 Then
        TEMP_MATRIX(i, 8) = DATA_MATRIX(i, 7) / DATA_MATRIX(i - 1, 7) - 1
    Else
        TEMP_MATRIX(i, 8) = DATA_MATRIX(i, 5) / DATA_MATRIX(i, 2) - 1
    End If
    
    TEMP_MATRIX(i, 9) = DATA_MATRIX(i, 6) / 1000
    TEMP_MATRIX(i, 10) = DATA_MATRIX(i, 7) / DATA_MATRIX(i, 5) * DATA_MATRIX(i, 2)
    TEMP_MATRIX(i, 11) = DATA_MATRIX(i, 7)
    
    If (i - DAYS_BACK) > 0 Then
        TEMP_MATRIX(i, 12) = TEMP_MATRIX(i, 11) / _
                             TEMP_MATRIX(i - DAYS_BACK, 10) - 1
        
        TEMP_MATRIX(i, 13) = TEMP_MATRIX(i - 1, 13)

        If (TEMP_MATRIX(i, 12) <= BUY_PERCENT And _
            TEMP_MATRIX(i - 1, 14) >= SHARES_TRADE * TEMP_MATRIX(i, 11)) Then
                    TEMP_MATRIX(i, 13) = SHARES_TRADE + TEMP_MATRIX(i, 13)
        Else
            If (TEMP_MATRIX(i, 12) >= SELL_PERCENT And _
                TEMP_MATRIX(i - 1, 13) >= SHARES_TRADE) Then
                    TEMP_MATRIX(i, 13) = -SHARES_TRADE + TEMP_MATRIX(i, 13)
'            Else
'                TEMP_MATRIX(i, 13) = 0 + TEMP_MATRIX(i, 13)
            End If
        End If
        
        If TEMP_MATRIX(i, 13) > TEMP_MATRIX(i - 1, 13) Then
            k = -1
        Else
            If TEMP_MATRIX(i, 13) < TEMP_MATRIX(i - 1, 13) Then
                k = 1
            Else
                k = 0
            End If
        End If
        
        Z_VAL = TEMP_MATRIX(i - 1, 14) + SHARES_TRADE * TEMP_MATRIX(i, 11) * k
        If Z_VAL > 0 Then
            TEMP_MATRIX(i, 14) = Z_VAL
        Else
            TEMP_MATRIX(i, 14) = 0
        End If
        
        
    Else
        TEMP_MATRIX(i, 12) = ""
        TEMP_MATRIX(i, 13) = INITIAL_SHARES
        TEMP_MATRIX(i, 14) = INITIAL_CASH
    End If
    
    TEMP_MATRIX(i, 15) = TEMP_MATRIX(i, 13) * TEMP_MATRIX(i, 11) + TEMP_MATRIX(i, 14)
    
    If i <> 1 Then
        TEMP_MATRIX(i, 16) = TEMP_MATRIX(i, 15) / TEMP_MATRIX(i - 1, 15) - 1
        MEAN_VAL = MEAN_VAL + TEMP_MATRIX(i, 16)
    Else
        TEMP_MATRIX(i, 16) = ""
    End If
Next i

If OUTPUT = 0 Then
    ASSET_DAYS_BACK_SIGNAL_FUNC = TEMP_MATRIX
    Exit Function
End If

MEAN_VAL = MEAN_VAL / (NROWS - 1)

For i = 2 To NROWS
    VOLAT_VAL = VOLAT_VAL + (TEMP_MATRIX(i, 16) - MEAN_VAL) ^ 2
Next i
VOLAT_VAL = (VOLAT_VAL / (NROWS - 1)) ^ 0.5

If OUTPUT = 1 Then
    ASSET_DAYS_BACK_SIGNAL_FUNC = MEAN_VAL / VOLAT_VAL
Else
    ASSET_DAYS_BACK_SIGNAL_FUNC = Array(MEAN_VAL / VOLAT_VAL, MEAN_VAL, VOLAT_VAL)
End If

Exit Function
ERROR_LABEL:
ASSET_DAYS_BACK_SIGNAL_FUNC = "--"
End Function
