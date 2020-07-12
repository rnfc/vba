Attribute VB_Name = "FINAN_ASSET_SYSTEM_FRACTIONAL"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
                    
'http://www.reuters.com/article/pressRelease/idUS122840+28-Jul-2009+BW20090728
'http://www.elitetrader.com/vb/showthread.php?threadid=170543
'http://www.gummy-stuff.org/kelly-ratio.htm
'http://www.trader-soft.com/money-management/optimal-f.html
'http://www.financialwebring.org/gummystuff/money-management.htm
'http://parametricplanet.com/rvince/

Function ASSET_FRACTIONAL_SIGNAL_FUNC(ByVal TICKER_STR As Variant, _
Optional ByVal INITIAL_DATE As Date, _
Optional ByVal END_DATE As Date, _
Optional ByRef BUY_PERCENT As Double = -0.03, _
Optional ByRef SELL_PERCENT As Double = 0.02, _
Optional ByVal INITIAL_SHARES As Long = 1000, _
Optional ByVal TRADE_PERCENT As Double = 0.05, _
Optional ByVal INITIAL_CASH As Double = 50000, _
Optional ByVal CASH_RATE As Double = 0.02, _
Optional ByVal COUNT_BASIS As Double = 365, _
Optional ByVal DELTA_VAL As Double = 0.010001, _
Optional ByVal OUTPUT As Integer = 0)

Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim ii As Long
Dim jj As Long

Dim SROW As Long
Dim NROWS As Long

Dim RATIO_VAL As Double
Dim FACTOR_VAL As Double

Dim MEAN_VAL As Double
Dim SIGMA_VAL As Double

Dim INDEX_ARR As Variant
Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

If INITIAL_SHARES = 0 Then: GoTo ERROR_LABEL
'------------------------------------------------------------------
If IsArray(TICKER_STR) = False Then
    DATA_MATRIX = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, _
                  INITIAL_DATE, END_DATE, "d", "DOHLCVA", False, _
                  True, True)
Else
    DATA_MATRIX = TICKER_STR
End If
'------------------------------------------------------------------
NROWS = UBound(DATA_MATRIX, 1)
'------------------------------------------------------------------

FACTOR_VAL = (1 + CASH_RATE) ^ (1 / COUNT_BASIS) - 1
NROWS = UBound(DATA_MATRIX, 1)
ReDim TEMP_MATRIX(0 To NROWS, 1 To 23) 'NROWS + 1

'------------------------------------------------------------------------------
TEMP_MATRIX(0, 1) = "DATE"
TEMP_MATRIX(0, 2) = "OPEN"
TEMP_MATRIX(0, 3) = "HIGH"
TEMP_MATRIX(0, 4) = "LOW"
TEMP_MATRIX(0, 5) = "CLOSE"
TEMP_MATRIX(0, 6) = "VOLUME"
TEMP_MATRIX(0, 7) = "ADJ.CLOSE"

TEMP_MATRIX(0, 8) = "RETURNS"
'------------------------------------------------------------------------------
TEMP_MATRIX(0, 9) = "BUY DOWN: " & Format(BUY_PERCENT, "0.00%")
TEMP_MATRIX(0, 10) = "SELL UP: " & Format(SELL_PERCENT, "0.00%")
TEMP_MATRIX(0, 11) = "SHARES HELD"
TEMP_MATRIX(0, 12) = "EQUITY"
TEMP_MATRIX(0, 13) = "CASH"
TEMP_MATRIX(0, 14) = "SYSTEM VALUE"
'------------------------------------------------------------------------------
TEMP_MATRIX(0, 15) = "SHARES TRADED"
TEMP_MATRIX(0, 16) = "SYSTEM SCALED"
TEMP_MATRIX(0, 17) = "SYSTEM VALUE"
TEMP_MATRIX(0, 18) = "PROFIT/LOSS"
TEMP_MATRIX(0, 19) = "SYSTEM MAX"
TEMP_MATRIX(0, 20) = "SYSTEM DRAWDOWN" 'Active Drawdown
TEMP_MATRIX(0, 21) = "BUY/HOLD"
TEMP_MATRIX(0, 22) = "BUY/HOLD MAX"
TEMP_MATRIX(0, 23) = "BUY/HOLD DRAWDOWN" 'Passive Drawdown
'------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
For j = 1 To 7: TEMP_MATRIX(1, j) = DATA_MATRIX(1, j): Next j
TEMP_MATRIX(1, 6) = TEMP_MATRIX(1, 6) / 1000

TEMP_MATRIX(1, 8) = 0
TEMP_MATRIX(1, 9) = 0
TEMP_MATRIX(1, 10) = 0
TEMP_MATRIX(1, 11) = INITIAL_SHARES
TEMP_MATRIX(1, 12) = TEMP_MATRIX(1, 11) * TEMP_MATRIX(1, 7)
TEMP_MATRIX(1, 13) = INITIAL_CASH
TEMP_MATRIX(1, 14) = TEMP_MATRIX(1, 13) + TEMP_MATRIX(1, 12)
    
TEMP_MATRIX(1, 15) = ""
TEMP_MATRIX(1, 16) = TEMP_MATRIX(1, 7)

TEMP_MATRIX(1, 17) = TEMP_MATRIX(1, 14)
TEMP_MATRIX(1, 18) = ""

TEMP_MATRIX(1, 19) = TEMP_MATRIX(1, 14)
RATIO_VAL = TEMP_MATRIX(1, 19) / TEMP_MATRIX(1, 7)

TEMP_MATRIX(1, 20) = TEMP_MATRIX(1, 19) - TEMP_MATRIX(1, 14)
TEMP_MATRIX(1, 21) = RATIO_VAL * TEMP_MATRIX(1, 7)
TEMP_MATRIX(1, 22) = RATIO_VAL * TEMP_MATRIX(1, 7)
TEMP_MATRIX(1, 23) = TEMP_MATRIX(1, 22) - RATIO_VAL * TEMP_MATRIX(1, 7)

'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------

k = 1
l = 0
ReDim INDEX_ARR(1 To 1)
MEAN_VAL = 0
For i = 2 To NROWS
    
    For j = 1 To 7: TEMP_MATRIX(i, j) = DATA_MATRIX(i, j): Next j
    TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 6) / 1000
     
     TEMP_MATRIX(i, 8) = TEMP_MATRIX(i, 7) / TEMP_MATRIX(i - 1, 7) - 1
    
    If TEMP_MATRIX(i, 8) < BUY_PERCENT Then
        TEMP_MATRIX(i, 9) = TEMP_MATRIX(i, 7)
    Else
        TEMP_MATRIX(i, 9) = 0
    End If
    
    If TEMP_MATRIX(i, 8) > SELL_PERCENT Then
        TEMP_MATRIX(i, 10) = TEMP_MATRIX(i, 7)
    Else
        TEMP_MATRIX(i, 10) = 0
    End If
        
    If TEMP_MATRIX(i, 9) > 0 Then
        ii = 1
    Else
        ii = 0
    End If
        
    If TEMP_MATRIX(i, 10) > 0 Then
        jj = 1
    Else
        jj = 0
    End If
        
    TEMP_MATRIX(i, 11) = Int(TEMP_MATRIX(i - 1, 11) * (1 + (TRADE_PERCENT * (ii - jj))))
        
    TEMP_MATRIX(i, 12) = TEMP_MATRIX(i, 11) * TEMP_MATRIX(i, 7)
        
    If TEMP_MATRIX(i, 11) > TEMP_MATRIX(i - 1, 11) Then
        ii = 1
    Else
        ii = 0
    End If
        
    If TEMP_MATRIX(i - 1, 11) > TEMP_MATRIX(i, 11) Then
        jj = 1
    Else
        jj = 0
    End If
        
    TEMP_MATRIX(i, 13) = TEMP_MATRIX(i - 1, 13) * (1 + FACTOR_VAL) - _
                            (TEMP_MATRIX(i, 11) - TEMP_MATRIX(i - 1, 11)) * _
                            TEMP_MATRIX(i, 9) * ii + (TEMP_MATRIX(i - 1, 11) _
                            - TEMP_MATRIX(i, 11)) * TEMP_MATRIX(i, 10) * jj
    
    TEMP_MATRIX(i, 14) = TEMP_MATRIX(i, 13) + TEMP_MATRIX(i, 12)
    MEAN_VAL = MEAN_VAL + (TEMP_MATRIX(i, 14) / TEMP_MATRIX(i - 1, 14) - 1)
    
    If TEMP_MATRIX(i, 11) > TEMP_MATRIX(i - 1, 11) Then
        TEMP_MATRIX(i, 15) = "Buy " & Format(TEMP_MATRIX(i, 11) - _
                                             TEMP_MATRIX(i - 1, 11), "0")
    Else
        If TEMP_MATRIX(i - 1, 11) > TEMP_MATRIX(i, 11) Then
            TEMP_MATRIX(i, 15) = "Sell " & Format(TEMP_MATRIX(i - 1, 11) - _
                                                  TEMP_MATRIX(i, 11), "0")
        Else
            TEMP_MATRIX(i, 15) = ""
        End If
    End If
    
    TEMP_MATRIX(i, 16) = TEMP_MATRIX(i - 1, 16) * TEMP_MATRIX(i, 14) / TEMP_MATRIX(i - 1, 14)
    If TEMP_MATRIX(i, 11) <> TEMP_MATRIX(i - 1, 11) Then
       TEMP_MATRIX(i, 17) = TEMP_MATRIX(i, 14)
       TEMP_MATRIX(i, 18) = TEMP_MATRIX(i, 14) - TEMP_MATRIX(k, 14)
       k = i
    Else
       TEMP_MATRIX(i, 17) = ""
       TEMP_MATRIX(i, 18) = ""
    End If

    If TEMP_MATRIX(i, 18) <> "" Then
        l = l + 1 'Number of trades
        ReDim Preserve INDEX_ARR(1 To l)
        INDEX_ARR(l) = TEMP_MATRIX(i, 18)
    End If
    If TEMP_MATRIX(i, 14) > TEMP_MATRIX(i - 1, 19) Then
        TEMP_MATRIX(i, 19) = TEMP_MATRIX(i, 14)
    Else
        TEMP_MATRIX(i, 19) = TEMP_MATRIX(i - 1, 19)
    End If
    TEMP_MATRIX(i, 20) = TEMP_MATRIX(i, 19) - TEMP_MATRIX(i, 14)
    TEMP_MATRIX(i, 21) = RATIO_VAL * TEMP_MATRIX(i, 7)
    If (RATIO_VAL * TEMP_MATRIX(i, 7)) > TEMP_MATRIX(i - 1, 22) Then
        TEMP_MATRIX(i, 22) = RATIO_VAL * TEMP_MATRIX(i, 7)
    Else
        TEMP_MATRIX(i, 22) = TEMP_MATRIX(i - 1, 22)
    End If
    TEMP_MATRIX(i, 23) = TEMP_MATRIX(i, 22) - RATIO_VAL * TEMP_MATRIX(i, 7)
Next i

INDEX_ARR = FIXED_FRACTIONAL_TRADING_SCHEME_OBJ_FUNC(INDEX_ARR, DELTA_VAL, 4)
SROW = LBound(INDEX_ARR)
'------------------------------------------------------------------------------
TEMP_MATRIX(0, 15) = "SHARES TRADED: " & Format(Int(INDEX_ARR(SROW) * TEMP_MATRIX(1, 22) / _
INDEX_ARR(SROW + 1)), "0") & " SHARES PER TRADE (N)"
TEMP_MATRIX(0, 18) = "PROFIT/LOSS: " & Format(INDEX_ARR(SROW), "0.00%") & " OPTIMAL F"
TEMP_MATRIX(0, 19) = "SYSTEM MAX: " & Format(l, "0") & " NUMBER OF TRADES"
TEMP_MATRIX(0, 20) = "SYSTEM DRAWDOWN: " & Format(INDEX_ARR(SROW + 1), "$0") & " MAXIMUM LOSS"
'------------------------------------------------------------------------------

If OUTPUT = 0 Then
    ASSET_FRACTIONAL_SIGNAL_FUNC = TEMP_MATRIX
    Exit Function
End If

MEAN_VAL = MEAN_VAL / (NROWS - 1)
SIGMA_VAL = 0
For i = 2 To NROWS
    SIGMA_VAL = SIGMA_VAL + ((TEMP_MATRIX(i, 14) / TEMP_MATRIX(i - 1, 14) - 1) - MEAN_VAL) ^ 2
Next i
SIGMA_VAL = (SIGMA_VAL / (NROWS - 1)) ^ 0.5

If OUTPUT = 1 Then
    ASSET_FRACTIONAL_SIGNAL_FUNC = MEAN_VAL / SIGMA_VAL
Else
    ASSET_FRACTIONAL_SIGNAL_FUNC = Array(MEAN_VAL / SIGMA_VAL, MEAN_VAL, SIGMA_VAL)
End If

Exit Function
ERROR_LABEL:
ASSET_FRACTIONAL_SIGNAL_FUNC = "--"
End Function

'If you expect a worst-case "possible" loss per share of $L and you trade N shares, you might
'expect a possible loss of $L*N (for an N-share trade).

'If you insist that this be no greater than a fraction f of your equity (worth $E), then you'd
'have: $L*N = $f*E. That means you should trade: N = f*E / L shares per trade.

'For example, if we are prepared to lose no more than L = $5 per trade and we have E = $30,000
'in stock, then we should trade N = f*E / L = f*30,000/5. Using a 5% fraction, so f = 0.05,
'we'd get: N = 300 shares per trade. For a $30 stock, that's 300*$30 = $9,000 for a single
'trade. Isn't that a bit much?

'It 's agressive, especially designed for them that don't mind risk.
'I understand that the most common figure is f = 0.02, or 2% of your Equity, so N = 0.02*30,000/5
'= 120 shares, worth $3600. I should also point out that you can consider a fraction f of your
'total portfolio (including cash) or a fraction f of your equity dollars or a fraction f of
'the number of shares or ...

'Yeah, I understand ... it's up to me. If I use N = f*E / L, I'd be increasing the number of shares
'traded as my Equity increased, right. Yes. If the stock is trading at $60 per share, and you put in
'a stop loss at $58 (so you'd automatically sell if the stock drops from $60 to $58), then you can
'lose no more than L = $2.00 per share.

'If you're Equity is $25,000 and you're using f = 0.02, or 2%, then your trade size should be no more
'than 0.02*25,000/2.00 = 250 shares. If the stock price doubled (and you kept L = $2.00 loss per share)
'then the trade size would double. So how do I know whether to do the 2% or 5% or ...
