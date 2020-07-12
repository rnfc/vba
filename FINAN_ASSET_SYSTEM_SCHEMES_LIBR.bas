Attribute VB_Name = "FINAN_ASSET_SYSTEM_SCHEMES_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : ASSET_BLOCK_SCHEME_FUNC
'DESCRIPTION   : Net Flow at best prevailing asset quote
'LIBRARY       : FINAN_ASSET_TRADE
'GROUP         : BLOCK
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 03/16/2009
'************************************************************************************
'************************************************************************************

Function ASSET_BLOCK_SCHEME_FUNC(ByRef ASSET1_PRICE_RNG As Variant, _
ByRef ASSET1_QUANTITY_RNG As Variant, _
ByRef ASSET2_PRICE_RNG As Variant, _
ByRef ASSET2_QUANTITY_RNG As Variant)

'To show prices use the 0.125 as a ratio (e.g. 18 1/8, 26 7/8)

Dim j As Long
Dim NCOLUMNS As Long

Dim ASSET1_PRICE As Double
Dim ASSET1_QUANTITY As Double
Dim ASSET2_PRICE As Double
Dim ASSET2_QUANTITY As Double

Dim ASSET1_PRICE_VECTOR As Variant
Dim ASSET1_QUANTITY_VECTOR As Variant
Dim ASSET2_PRICE_VECTOR As Variant
Dim ASSET2_QUANTITY_VECTOR As Variant

Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

'-----------------------------------------------------------------------------------
If IsArray(ASSET1_PRICE_RNG) = True Then
    ASSET1_PRICE_VECTOR = ASSET1_PRICE_RNG
    If UBound(ASSET1_PRICE_VECTOR, 2) = 1 Then
        ASSET1_PRICE_VECTOR = MATRIX_TRANSPOSE_FUNC(ASSET1_PRICE_VECTOR)
    End If
Else
    ReDim ASSET1_PRICE_VECTOR(1 To 1, 1 To 1)
    ASSET1_PRICE_VECTOR(1, 1) = ASSET1_PRICE_RNG
End If
NCOLUMNS = UBound(ASSET1_PRICE_VECTOR, 2)
'-----------------------------------------------------------------------------------
If IsArray(ASSET1_QUANTITY_RNG) = True Then
    ASSET1_QUANTITY_VECTOR = ASSET1_QUANTITY_RNG
    If UBound(ASSET1_QUANTITY_VECTOR, 2) = 1 Then
        ASSET1_QUANTITY_VECTOR = MATRIX_TRANSPOSE_FUNC(ASSET1_QUANTITY_VECTOR)
    End If
Else
    ReDim ASSET1_QUANTITY_VECTOR(1 To 1, 1 To 1)
    ASSET1_QUANTITY_VECTOR(1, 1) = ASSET1_QUANTITY_RNG
End If
If NCOLUMNS <> UBound(ASSET1_QUANTITY_VECTOR, 2) Then: GoTo ERROR_LABEL
'-----------------------------------------------------------------------------------
If IsArray(ASSET2_PRICE_RNG) = True Then
    ASSET2_PRICE_VECTOR = ASSET2_PRICE_RNG
    If UBound(ASSET2_PRICE_VECTOR, 2) = 1 Then
        ASSET2_PRICE_VECTOR = MATRIX_TRANSPOSE_FUNC(ASSET2_PRICE_VECTOR)
    End If
Else
    ReDim ASSET2_PRICE_VECTOR(1 To 1, 1 To 1)
    ASSET2_PRICE_VECTOR(1, 1) = ASSET2_PRICE_RNG
End If
If NCOLUMNS <> UBound(ASSET2_PRICE_VECTOR, 2) Then: GoTo ERROR_LABEL
'-----------------------------------------------------------------------------------
If IsArray(ASSET2_QUANTITY_RNG) = True Then
    ASSET2_QUANTITY_VECTOR = ASSET2_QUANTITY_RNG
    If UBound(ASSET2_QUANTITY_VECTOR, 2) = 1 Then
        ASSET2_QUANTITY_VECTOR = MATRIX_TRANSPOSE_FUNC(ASSET2_QUANTITY_VECTOR)
    End If
Else
    ReDim ASSET2_QUANTITY_VECTOR(1 To 1, 1 To 1)
    ASSET2_QUANTITY_VECTOR(1, 1) = ASSET2_QUANTITY_RNG
End If
If NCOLUMNS <> UBound(ASSET2_QUANTITY_VECTOR, 2) Then: GoTo ERROR_LABEL
'-----------------------------------------------------------------------------------
ReDim TEMP_MATRIX(1 To 3, 1 To NCOLUMNS + 1)
TEMP_MATRIX(1, 1) = "NET FLOW AT BEST PREVAILING QUOTES:"
TEMP_MATRIX(2, 1) = "BLOCK DEAL:"
TEMP_MATRIX(3, 1) = "RATIO:"
'-----------------------------------------------------------------------------------

'-----------------------------------------------------------------------------------
For j = 1 To NCOLUMNS
'-----------------------------------------------------------------------------------
    ASSET1_PRICE = ASSET1_PRICE_VECTOR(1, j)
    ASSET1_QUANTITY = ASSET1_QUANTITY_VECTOR(1, j)
    ASSET2_PRICE = ASSET2_PRICE_VECTOR(1, j)
    ASSET2_QUANTITY = ASSET2_QUANTITY_VECTOR(1, j)
    TEMP_MATRIX(1, j + 1) = ASSET1_PRICE * ASSET1_QUANTITY - ASSET2_PRICE * ASSET2_QUANTITY
    TEMP_MATRIX(2, j + 1) = ASSET1_PRICE * ASSET1_QUANTITY + ASSET2_PRICE * ASSET2_QUANTITY
    TEMP_MATRIX(3, j + 1) = TEMP_MATRIX(1, j + 1) / TEMP_MATRIX(2, j + 1)
'-----------------------------------------------------------------------------------
Next j
'-----------------------------------------------------------------------------------

ASSET_BLOCK_SCHEME_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
ASSET_BLOCK_SCHEME_FUNC = Err.number
End Function


'http://www.google.ca/search?hl=en&q=fixed+ratio+%22ryan+jones%22&btnG=Search&meta=

Function FIXED_RATIO_TRADING_SCHEME_FUNC(ByVal P_VAL As Double, _
ByVal D_VAL As Double)

'N = (1/2) x (1 + (1+8P / delta)^0.5)
'where
'P is the expected profit per trade (example: P = $1250)

'delta is the per-contract profit required to increase you holdings by one contract
'(example: delta = $1000)

'and N is the number of contracts devoted to each trade (example:
'N = (1/2)[ 1 + (1+8*1250/1000)^0.5 ]) = 2.2

On Error GoTo ERROR_LABEL

FIXED_RATIO_TRADING_SCHEME_FUNC = (1 / 2) * (1 + (1 + 8 * P_VAL / D_VAL) ^ 0.5)

Exit Function
ERROR_LABEL:
FIXED_RATIO_TRADING_SCHEME_FUNC = Err.number
End Function



'If you expect a worst-case "possible" loss per share of $L and you trade N
'shares, you might expect a possible loss of $L*N (for an N-share trade).

'If you insist that this be no greater than a fraction f of your equity (worth $E),
'then you'd have: $L*N = $f*E.

'That means you should trade: N = f*E / L shares per trade.

'For example, if we are prepared to lose no more than L = $5 per trade and we have
'E = $30,000 in stock, then we should trade N = f*E / L = f*30,000/5.
'Using a 5% fraction, so f = 0.05, we'd get: N = 300 shares per trade.

Function FIXED_FRACTIONAL_TRADING_SCHEME_FUNC(ByVal F_VAL As Double, _
ByVal E_VAL As Double, _
ByVal L_VAL As Double)

'N = f * E / L
'where
'f is some fraction of your holdings (example: f = 0.02, meaning 2%)
'E is the dollar amount of your holdings (example: E = $50,000)
'L is the maximum loss per share (example: L = $10)
'and N is the number of shares devoted to each trade (example: N =
'0.02*50000/10 = 100 shares ... or contracts)

On Error GoTo ERROR_LABEL

FIXED_FRACTIONAL_TRADING_SCHEME_FUNC = F_VAL * E_VAL / L_VAL

Exit Function
ERROR_LABEL:
FIXED_FRACTIONAL_TRADING_SCHEME_FUNC = Err.number
End Function

'------------------------------------------------------------------------------------------------
'Optimal f Trading Scheme:
'------------------------------------------------------------------------------------------------
'Calculate the value of f that maximizes: W = (1 + f T1/L)(1 + f T2/L)...(1 + f Tn/L)
'where

'Tk is the trading profit or loss for the kth trade (example: T1 = $123, T2 = -$456, etc.)
'L is the maximum loss per trade (example: L = $100)

'If the "optimal" f is denoted by fopt (example: f = 0.30)
'and your Equity is $E (example: E = $30,000)

'N = fopt E / L   is the number of shares or contracts per trade.
'(example: N = (0.30)($30,000)/$100 = 90 contracts)

'Note that E / L gives the number of "Maximum Losses" that'll wipe our your Equity.
'Note, too, that L, the largest LOSS, is the largest negative value of T1, T2, ... Tn.

'So if fopt turns out to be, say 60%, something like the picture above, then I risk 60% of my
'Equity, right?
'Good heavens! NO!! That's just the preamble to "optimal f". Your trade should involve
'N = fopt E / L shares.

'Note that N*L is what you'd lose if every trade lost the maximum amount: $L. That maximum total
'loss, as a fraction of your Equity, is fopt.

'For example, if your Equity was E = $25,000, then with fopt = 0.60, you'd trade
'N = (0.60)(25,000)/(85.64) = 175 shares (or contracts)

'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
'Reference: Ralph Vince in the book "Portfolio Management Formulas : Mathematical Trading
'Methods for the Futures, Options, and Stock Markets".
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------

Function OPTIMAL_FIXED_FRACTIONAL_TRADING_SCHEME_FUNC(ByRef PROFIT_LOSSES_RNG As Variant, _
Optional ByVal DELTA_VAL As Double = 0.01)

Dim i As Long
Dim NROWS As Long

Dim F_VAL As Double
Dim W_VAL As Double
Dim MAX_VAL As Double
Dim F_OPT_VAL As Double
Dim MIN_VAL As Double
Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = PROFIT_LOSSES_RNG ' per trade
NROWS = UBound(DATA_VECTOR, 1)

MIN_VAL = 2 ^ 52
For i = 1 To NROWS
    If DATA_VECTOR(i, 1) < MIN_VAL Then: MIN_VAL = DATA_VECTOR(i, 1)
Next i
MIN_VAL = MIN_VAL * -1

MAX_VAL = -1
For F_VAL = 0 To 1 Step DELTA_VAL
    W_VAL = 1
    For i = 1 To NROWS 'trade #
       W_VAL = W_VAL * (1 + F_VAL * DATA_VECTOR(i, 1) / MIN_VAL)
    Next i
    If W_VAL > MAX_VAL Then
        MAX_VAL = W_VAL
        F_OPT_VAL = F_VAL
    End If
Next F_VAL

OPTIMAL_FIXED_FRACTIONAL_TRADING_SCHEME_FUNC = F_OPT_VAL

'If you lose $L with each trade and you've got $E in your portfolio, then E / L
'trades later ... you're dead.

'Yeah, I get that, but what's that optimal f thing?
'Suppose you make n trades and the n returns, as a percentage of your Equity, are: r1, r2 ... rn.
'Then, after n trades, $1.00 will grow to W = $(1+r1)(1+r2)...(1+rn).
'The profit or loss associated with your 1st trade is T1 = r1E1 where E1 is your Equity just
'before the trade ... with similar equations for subsequent trades.

'Then the per-trade returns can be written as: rk = Tk / Ek   ... with k going from 1 to n.
'The gain factor, after n trades, can now be written as: W = (1+T1 / E1) (1+T2 / E2)... (1+Tn / En)
'where your profits/losses per trade are T1, T2 ... Tn with some positive (that's profit) and
'some negative (that a loss).

'Denote the largest loss by $L ... taken as a positive number. That is: L = - min [ T1, T2 ... Tn]
'If we regard L as our "risk", then what fraction of our current Equity is that maximum loss?
'It's f = L / Ek.

'Okay, so we want to trade just that many shares (or contracts) on each trade so that:
'f = L / Ek is a constant.

'With this definition of f, we see that: Tk / Ek = f Tk / L
'Then our n-trade gain factor becomes:   W = (1 + f T1/L)(1 + f T2/L)...(1 + f Tn/L)
'... and that's what we want to maximize by clever choice of f.

'Yeah, but are you sure there is a maximum ... and how do you find it?
'if you insist that f lie between 0 and 1, then you will "usually" have a maximum as shown here:
'In this example, the maximum loss (at Trade #12) is L = $63.39 and the gain factor that we're
'maximizing is: W = ( 1+(65.13) f / (63.19) ) ( 1+(-41.33) f / (63.19) )( 1+(48.68) f / (63.19) ) ...
'( 1+(-45.76) f / (63.19) )

Exit Function
ERROR_LABEL:
OPTIMAL_FIXED_FRACTIONAL_TRADING_SCHEME_FUNC = Err.number
End Function


Function FIXED_FRACTIONAL_TRADING_SCHEME_OBJ_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal DELTA_VAL As Double = 0.05, _
Optional ByVal OUTPUT As Integer = 1)

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim F_VAL As Double
Dim W_VAL As Double

Dim TEMP_MIN As Double
Dim TEMP_MAX As Double
Dim TEMP_VAL As Double

Dim TEMP_MATRIX As Variant

Dim tolerance As Double

On Error GoTo ERROR_LABEL

tolerance = 1E+15
NROWS = UBound(DATA_RNG)
NCOLUMNS = (1 - 0) / DELTA_VAL

ReDim TEMP_MATRIX(0 To NROWS, 0 To NCOLUMNS)
TEMP_MATRIX(0, 0) = "F/W"

TEMP_MIN = tolerance
For i = 1 To NROWS
    If DATA_RNG(i) <= TEMP_MIN Then: TEMP_MIN = DATA_RNG(i)
    TEMP_MATRIX(i, 0) = DATA_RNG(i)
Next i
TEMP_MIN = TEMP_MIN * -1 'Maximum Loss

F_VAL = 0
TEMP_MAX = -1
For j = 1 To NCOLUMNS
    TEMP_MATRIX(0, j) = F_VAL
    W_VAL = 1
    For i = 1 To NROWS
       W_VAL = W_VAL * (1 + F_VAL * DATA_RNG(i) / TEMP_MIN)
       TEMP_MATRIX(i, j) = W_VAL
    Next i
    If W_VAL > TEMP_MAX Then
        TEMP_MAX = W_VAL
        TEMP_VAL = F_VAL
    End If
    F_VAL = F_VAL + DELTA_VAL
Next j

'------------------------------------------------------------------------------------
Select Case OUTPUT
'------------------------------------------------------------------------------------
    Case 0 'Optimal Factor
'------------------------------------------------------------------------------------
        FIXED_FRACTIONAL_TRADING_SCHEME_OBJ_FUNC = TEMP_VAL
'------------------------------------------------------------------------------------
    Case 1 'Maximum Loss
'------------------------------------------------------------------------------------
        FIXED_FRACTIONAL_TRADING_SCHEME_OBJ_FUNC = TEMP_MIN
'------------------------------------------------------------------------------------
    Case 2 'Optimal Factor, Maximum Loss
'------------------------------------------------------------------------------------
        FIXED_FRACTIONAL_TRADING_SCHEME_OBJ_FUNC = Array(TEMP_VAL, TEMP_MIN)
'------------------------------------------------------------------------------------
    Case 3
'------------------------------------------------------------------------------------
        FIXED_FRACTIONAL_TRADING_SCHEME_OBJ_FUNC = TEMP_MATRIX
'------------------------------------------------------------------------------------
    Case Else
'------------------------------------------------------------------------------------
        FIXED_FRACTIONAL_TRADING_SCHEME_OBJ_FUNC = Array(TEMP_VAL, TEMP_MIN, TEMP_MATRIX)
'------------------------------------------------------------------------------------
End Select
'------------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
FIXED_FRACTIONAL_TRADING_SCHEME_OBJ_FUNC = Err.number
End Function




'plot of W versus f so you can see when (if?) there's an f-value in 0 < f < 1 that provides a maximum.

Function PLOT_FIXED_FRACTIONAL_TRADING_SCHEME_FUNC(ByRef PROFIT_LOSSES_RNG As Variant, _
Optional ByVal MIN_BIN_VAL As Double = 0, _
Optional ByVal MAX_BIN_VAL As Double = 1, _
Optional ByVal DELTA_BIN_VAL As Double = 0.05)

Dim i As Long
Dim j As Long
Dim NROWS As Long
Dim NCOLUMNS As Long
Dim MIN_VAL As Double
Dim TEMP_VAL As Double
Dim DATA_VECTOR As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = PROFIT_LOSSES_RNG ' per trade
NROWS = UBound(DATA_VECTOR, 1)
NCOLUMNS = Int((MAX_BIN_VAL - MIN_BIN_VAL) / DELTA_BIN_VAL) + IIf(MIN_BIN_VAL = 0, 2, 1)

ReDim TEMP_MATRIX(1 To NROWS + 3, 1 To NCOLUMNS + 2)

TEMP_MATRIX(1, 1) = "f"
TEMP_MATRIX(1, 2) = ":"

TEMP_MATRIX(2, 1) = "W"
TEMP_MATRIX(2, 2) = ":"

TEMP_MATRIX(3, 1) = "n"
TEMP_MATRIX(3, 2) = ":"

MIN_VAL = 2 ^ 52
For i = 1 To NROWS
    If DATA_VECTOR(i, 1) < MIN_VAL Then: MIN_VAL = DATA_VECTOR(i, 1)
    TEMP_MATRIX(i + 3, 1) = i
    TEMP_MATRIX(i + 3, 2) = DATA_VECTOR(i, 1)
Next i
MIN_VAL = MIN_VAL * -1

TEMP_VAL = MIN_BIN_VAL
For j = 1 To NCOLUMNS
    TEMP_MATRIX(1, j + 2) = TEMP_VAL
    TEMP_MATRIX(3, j + 2) = "--"
    i = 1
    TEMP_MATRIX(i + 3, j + 2) = (1 + TEMP_MATRIX(1, j + 2) * TEMP_MATRIX(i + 3, 2) / MIN_VAL)
    For i = 2 To NROWS
        TEMP_MATRIX(i + 3, j + 2) = TEMP_MATRIX(i + 2, j + 2) * (1 + TEMP_MATRIX(1, j + 2) * TEMP_MATRIX(i + 3, 2) / MIN_VAL)
    Next i
    TEMP_VAL = TEMP_VAL + DELTA_BIN_VAL
Next j

For j = 1 To NCOLUMNS
    TEMP_MATRIX(2, j + 2) = TEMP_MATRIX(NROWS + 3, j + 2)
Next j

PLOT_FIXED_FRACTIONAL_TRADING_SCHEME_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
PLOT_FIXED_FRACTIONAL_TRADING_SCHEME_FUNC = Err.number
End Function
