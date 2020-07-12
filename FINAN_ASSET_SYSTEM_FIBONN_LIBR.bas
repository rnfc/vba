Attribute VB_Name = "FINAN_ASSET_SYSTEM_FIBONN_LIBR"

'-------------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------------
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'-------------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------------
Private PUB_DATA_MATRIX As Variant
Private PUB_REFERENCE_DATE As Date
Private PUB_FIBONACCI_FLAG As Boolean
Private PUB_INITIAL_CASH As Double
'-------------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------------

Function ASSET_BOLLINGER_FIBONACCI_SIGNAL_FUNC(ByRef TICKER_STR As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal END_DATE As Date, _
Optional ByVal REFERENCE_DATE As Date, _
Optional ByVal BOLLINGER_SD As Double = 2, _
Optional ByVal BOLLINGER_MA_PERIODS As Long = 20, _
Optional ByVal SELL_BELOW As Double = -0.04, _
Optional ByVal BUY_BELOW As Double = 0.04, _
Optional ByVal MA1_PERIODS As Long = 30, _
Optional ByVal MA2_PERIODS As Long = 5, _
Optional ByVal FIBONACCI_FLAG As Boolean = True, _
Optional ByVal INITIAL_CASH As Double = 1000, _
Optional ByVal OUTPUT As Integer = 0)

'FIBONACCI_FLAG: Want an UP Fibonacci fan?

Dim g As Long
Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long
Dim m As Long
Dim n As Long

Dim o As Long
Dim p As Long

Dim ii As Long
Dim jj As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double
Dim SLOPE_VAL As Double

Dim MEAN_VAL As Double
Dim SIGMA_VAL As Double

Dim ATEMP_SUM As Double
Dim BTEMP_SUM As Double
Dim CTEMP_SUM As Double
Dim DTEMP_SUM As Double
Dim ETEMP_SUM As Double
Dim FTEMP_SUM As Double
Dim GTEMP_SUM As Double

'-------------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------------
'Following Mr. Bollinger, we do this:

'1. Plot the price of a stock, each day, for the past umpteen months.
'2. Pick a number like N = 20, then, for each day, compute the Standard Deviation
'and Average for the previous N days.

'3. Pick a small number like k = 2, then, for each day, compute
'Upper Bollinger = Average + k (Standard Deviation) * and
'Lower Bollinger = Average - k (Standard Deviation).

'4. Plot Upper Bollinger and Lower Bollinger along with the stock Price.
'* Standard Deviation2 = SD2 = (1/N)? (Pk - A)2
'where there are N stock prices, Pk (k=1 to N), and A = (1/N)?Pk is their average and
'SD is the Mean Square Deviation between the prices and their average and it can also be
'computed like so: SD2 = (1/N)? Pk2 - {(1/N)?Pk}2 namely the difference between the average
'of the squares and the square of the average ... nice, eh?

'where the Price of the stock seems to bounce between the Bollinger bands.
'(I should mention that Mr. Bollinger picks N=20 and k=2, but we can pick any numbers ... right?)
'Anyway, (if you have a great imagination) you might think that when the stock Price crashes thru' the
'Upper Bolli-band, we should SELL, and when it drops below the Lower Bolli-band we should BUY.

'However, when it crashes thru' the Upper-B it may keep going and who'd want to sell then? So maybe
'we wait for it go above then drop back below the Upper-B ... then we SELL.

'Now some folk wait for another piece of data to indicate a SELL (besides going above then below
'the Upper-B). That's the (are you ready for this) Relative Strength Index (or RSI) which measures
'the percentage of times when the stock Price increased over the past N days.

'Because it 's a percentage it goes from 0 to 100 and when the stock Price goes above then below
'the Upper-B (that's part of our SELL signal) and, in addition, the RSI is at least 70% (meaning
'the stock increased at least 70% of the time over the past N days), then we conclude that the stock
'has been overbought and we should SELL.

'Uh ... it's not really we who conclude but those who play with Bolli-bands.

'Anyway, we can also consider a BUY signal to be when the stock Price drops below the Lower-B then
'rises above again ... AND the RSI is less than 30% (meaning the number of increases over the past
'N days is no more than 30%).

'Of course, the 70% and 30% are arbitrary.
'Here 's a graph of what the RSI might look like ... and the levels 60% and 40%

'For example, things might look like so:

'If Bolli-bands don't improve your financial health you might want to try band-aids ... and/or Tylenol.

'P.S.   I have it on good authority that Bollinger himself does NOT think that using Bollinger bands
'and RSI as described above is a good idea. Indeed, John Bollinger's words are:
'"Perhaps you would be so kind as to mention that John Bollinger, the father of the eponymous bands,
'thinks that using Bollinger Bands and RSI in the manner described here, is a pretty poor idea."

'John Bollinger, CFA, CMT
'-------------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------------
Dim GOLDEN_RATIOS_VAL As Double

'Once upon a time, an Italian mathematician called Leonardo Fibonacci (while studying the
'population growth in rabbits) considered the sequence of numbers: 1, 1, 2, 3, 5, 8, 13, ...
'where each number is the sum of the two previous numbers (so, for example, the next number
'would be 13 + 8 = 21).

'The numbers satisfy the equation Fn+2 = Fn+1 + Fn with F1 = F2 = 1.

'The ratio of successive numbers satisfies Fn+2/Fn+1 = 1 + 1/{Fn+1/Fn} and if we let n become
'infinite we get the limiting value of this ratio, namely x which satisfies:
'x = 1 + 1/x   or   x2 - x - 1 = 0   which has as a solution
'x = (1/2)[1 + SQRT(5)] = 1.618.
'Notice that 1/x = x - 1   so   1/1.618 = 0.618 (or 61.8%).
'Notice that 1 - 1/x = 1 - 0.618 (or 38.2%).
'The number x is called the Golden Ratio. (See Golden Ratios.)
'An interesting note: Divide a line into two parts so as to have these Ratios:
'x / 1 = (1 - x) / x then (surprise!) x = 0.618

'In any case, this number has been applied to so many things that it seemed inevitable that
'it'd be applied to the stock market. We'll talk about Fibonacci fans. To see a DOWN Fibonacci
'fan we do this:
'Draw a line from a Maximum stock price to a subsequent Minimum.
'This gives a Trend Line ... with some magic slope.
'We draw other lines (fans?) with slopes which are 61.8% and 38.2% of the Trend Line slope.
'Where these Fibonacci fans intersect the stock price chart we get (maybe) resistance levels or
'(maybe) buy & sell signals.....

'Reference: http://www.mcs.surrey.ac.uk/Personal/R.Knott/Fibonacci/phi.html
'-------------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------------

Dim DATA_MATRIX As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

If IsArray(TICKER_STR) = True Then
    DATA_MATRIX = TICKER_STR
Else
    DATA_MATRIX = _
    YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, START_DATE, END_DATE, _
    "d", "DOHLCVA", False, True, True)
End If
NROWS = UBound(DATA_MATRIX, 1)


If REFERENCE_DATE = 0 Or REFERENCE_DATE <= DATA_MATRIX(1, 1) Then
    REFERENCE_DATE = DATA_MATRIX(2, 1)
End If

NCOLUMNS = 24
ReDim TEMP_MATRIX(0 To NROWS, 1 To NCOLUMNS)
For j = 1 To NCOLUMNS: For i = 1 To NROWS: TEMP_MATRIX(i, j) = "": Next i: Next j

TEMP_MATRIX(0, 1) = "DATE"
TEMP_MATRIX(0, 2) = "OPEN"
TEMP_MATRIX(0, 3) = "HIGH"
TEMP_MATRIX(0, 4) = "LOW"
TEMP_MATRIX(0, 5) = "CLOSE"
TEMP_MATRIX(0, 6) = "VOLUME"
TEMP_MATRIX(0, 7) = "ADJ CLOSE"
TEMP_MATRIX(0, 9) = "VOLUME(K) & PRICE"
TEMP_MATRIX(0, 8) = "VOLUME-WEIGHTED"
TEMP_MATRIX(0, 10) = "VOLATILITY"
TEMP_MATRIX(0, 11) = "UPPER " & BOLLINGER_MA_PERIODS & "-DAY BOLLI BANDS @ " & _
                     Format(BOLLINGER_SD, "0.0") & " SD"
TEMP_MATRIX(0, 12) = "LOWER " & BOLLINGER_MA_PERIODS & "-DAY BOLLI BANDS @ " & _
                     Format(BOLLINGER_SD, "0.0") & " SD"
TEMP_MATRIX(0, 14) = "MA " & MA1_PERIODS & " DAY"
TEMP_MATRIX(0, 15) = "MA " & MA2_PERIODS & " DAY"
TEMP_MATRIX(0, 17) = "SELL BELOW " & Format(SELL_BELOW, "0.00%")
TEMP_MATRIX(0, 18) = "BUY BELOW " & Format(BUY_BELOW, "0.00%")

TEMP_MATRIX(0, 22) = "INVESTED"
TEMP_MATRIX(0, 23) = "CASH"
TEMP_MATRIX(0, 24) = "PORTFOLIO"

'-----------------------------------------------------------------------------------
GOLDEN_RATIOS_VAL = 2 / (1 + Sqr(5))
k = 0: l = 0: m = 0: n = 0: o = 0: p = 0
'-----------------------------------------------------------------------------------
ATEMP_SUM = 0
BTEMP_SUM = 0
'-----------------------------------------------------------------------------------
CTEMP_SUM = 0
DTEMP_SUM = 0
'-----------------------------------------------------------------------------------
ETEMP_SUM = 0
FTEMP_SUM = 0
'-----------------------------------------------------------------------------------
GTEMP_SUM = 0
'-----------------------------------------------------------------------------------
For i = 1 To NROWS
'-----------------------------------------------------------------------------------
    TEMP_MATRIX(i, 1) = DATA_MATRIX(i, 1)
    TEMP_MATRIX(i, 2) = DATA_MATRIX(i, 2) 'Open
    TEMP_MATRIX(i, 3) = DATA_MATRIX(i, 3) 'High
    TEMP_MATRIX(i, 4) = DATA_MATRIX(i, 4) 'Low
    TEMP_MATRIX(i, 5) = DATA_MATRIX(i, 5) 'Close
    TEMP_MATRIX(i, 6) = DATA_MATRIX(i, 6) / 1000 'Volume
    TEMP_MATRIX(i, 7) = DATA_MATRIX(i, 7) 'Adj Close
    
    If TEMP_MATRIX(i, 1) >= REFERENCE_DATE Then
        ATEMP_SUM = ATEMP_SUM + TEMP_MATRIX(i, 6)
        BTEMP_SUM = BTEMP_SUM + TEMP_MATRIX(i, 6) * TEMP_MATRIX(i, 5)
        TEMP_MATRIX(i, 8) = ATEMP_SUM
        TEMP_MATRIX(i, 9) = BTEMP_SUM / ATEMP_SUM
        
        l = l + 1
        
        If TEMP_MATRIX(i, 1) > REFERENCE_DATE Then
            
            If l > BOLLINGER_MA_PERIODS + 1 Then
                k = i - BOLLINGER_MA_PERIODS
                CTEMP_SUM = CTEMP_SUM + TEMP_MATRIX(i, 5)
                CTEMP_SUM = CTEMP_SUM - TEMP_MATRIX(k - 1, 5)
                MEAN_VAL = CTEMP_SUM / (BOLLINGER_MA_PERIODS + 1)

                DTEMP_SUM = 0
                For h = i To k Step -1
                    DTEMP_SUM = DTEMP_SUM + (TEMP_MATRIX(h, 5) - MEAN_VAL) ^ 2
                Next h
                SIGMA_VAL = (DTEMP_SUM / (BOLLINGER_MA_PERIODS + 1)) ^ 0.5
                
                MEAN_VAL = (CTEMP_SUM + TEMP_MATRIX(h, 5)) / (BOLLINGER_MA_PERIODS + 2) 'Adjustment
            Else
                CTEMP_SUM = CTEMP_SUM + TEMP_MATRIX(i, 5)
                MEAN_VAL = CTEMP_SUM / l
                DTEMP_SUM = 0
                For h = i To g Step -1
                    DTEMP_SUM = DTEMP_SUM + (TEMP_MATRIX(h, 5) - MEAN_VAL) ^ 2
                Next h
                SIGMA_VAL = (DTEMP_SUM / l) ^ 0.5
            End If
        
            If TEMP_MATRIX(i - 1, 1) = REFERENCE_DATE Then
                TEMP_MATRIX(i, 13) = 0
                n = 0
            ElseIf TEMP_MATRIX(i - 1, 1) > REFERENCE_DATE Then
                TEMP_MATRIX(i, 13) = DATA_MATRIX(i - 1, 5) / DATA_MATRIX(m, 5)

                TEMP_MATRIX(i, 14) = 0
                TEMP_MATRIX(i, 15) = 0
                TEMP_MATRIX(i, 16) = 0

                If TEMP_MATRIX(i - 2, 1) > REFERENCE_DATE Then
                    n = n + 1
                    ETEMP_SUM = ETEMP_SUM + TEMP_MATRIX(i - 1, 13)
                    FTEMP_SUM = FTEMP_SUM + TEMP_MATRIX(i - 1, 13)
                    
                    If n < MA2_PERIODS Then
                        TEMP_MATRIX(i, 15) = FTEMP_SUM / n
                    Else
                        k = i - MA2_PERIODS
                        FTEMP_SUM = FTEMP_SUM - TEMP_MATRIX(k, 13)
                        TEMP_MATRIX(i, 15) = FTEMP_SUM / (MA2_PERIODS - 1)
                    End If
                    
                    If n < MA1_PERIODS Then
                        TEMP_MATRIX(i, 14) = ETEMP_SUM / n
                    Else
                        k = i - MA1_PERIODS
                        ETEMP_SUM = ETEMP_SUM - TEMP_MATRIX(k, 13)
                        TEMP_MATRIX(i, 14) = ETEMP_SUM / (MA1_PERIODS - 1)
                    End If

                    TEMP_MATRIX(i, 16) = TEMP_MATRIX(i, 14) - TEMP_MATRIX(i, 15)
                    
                    TEMP_MATRIX(i, 17) = IIf((TEMP_MATRIX(i - 1, 16) > SELL_BELOW And _
                                              TEMP_MATRIX(i, 16) < SELL_BELOW), _
                                              TEMP_MATRIX(i, 13), -1)
                    
                    TEMP_MATRIX(i, 18) = IIf((TEMP_MATRIX(i - 1, 16) < BUY_BELOW And _
                                              TEMP_MATRIX(i, 16) > BUY_BELOW), _
                                              TEMP_MATRIX(i, 13), -1)


                    If TEMP_MATRIX(i, 17) > 0 Then
                        TEMP_MATRIX(i, 22) = 0
                    Else
                        If TEMP_MATRIX(i, 18) > 0 And TEMP_MATRIX(i - 1, 23) > 0 Then
                            TEMP_MATRIX(i, 22) = TEMP_MATRIX(i - 1, 23)
                        Else
                            TEMP_MATRIX(i, 22) = TEMP_MATRIX(i - 1, 22) * TEMP_MATRIX(i, 7) / TEMP_MATRIX(i - 1, 7)
                        End If
                    End If
                    
                    If TEMP_MATRIX(i, 17) > 0 And TEMP_MATRIX(i - 1, 22) > 0 Then
                        TEMP_MATRIX(i, 23) = TEMP_MATRIX(i - 1, 22) * TEMP_MATRIX(i, 7) / TEMP_MATRIX(i - 1, 7)
                    Else
                        If TEMP_MATRIX(i, 18) > 0 Then
                            TEMP_MATRIX(i, 23) = 0
                        Else
                            TEMP_MATRIX(i, 23) = TEMP_MATRIX(i - 1, 23)
                        End If
                    End If
                    
                    TEMP_MATRIX(i, 24) = TEMP_MATRIX(i, 22) + TEMP_MATRIX(i, 23)
                    GTEMP_SUM = GTEMP_SUM + (TEMP_MATRIX(i, 24) / TEMP_MATRIX(i - 1, 24) - 1)
                    p = p + 1
                ElseIf TEMP_MATRIX(i - 2, 1) = REFERENCE_DATE Then
                    TEMP_MATRIX(i, 22) = 0
                    TEMP_MATRIX(i, 23) = INITIAL_CASH
                    TEMP_MATRIX(i, 24) = TEMP_MATRIX(i, 22) + TEMP_MATRIX(i, 23)
                    o = i + 1
                End If
            End If
        ElseIf TEMP_MATRIX(i, 1) = REFERENCE_DATE Then
            g = i
            m = g + BOLLINGER_MA_PERIODS - 1
            CTEMP_SUM = TEMP_MATRIX(i, 5)
            MEAN_VAL = CTEMP_SUM / l
            DTEMP_SUM = 0
            SIGMA_VAL = DTEMP_SUM
        End If
        
        TEMP_MATRIX(i, 10) = SIGMA_VAL
        TEMP_MATRIX(i, 11) = MEAN_VAL + SIGMA_VAL * BOLLINGER_SD
        TEMP_MATRIX(i, 12) = MEAN_VAL - SIGMA_VAL * BOLLINGER_SD
        
    End If
'-----------------------------------------------------------------------------------
Next i
'-----------------------------------------------------------------------------------
TEMP_MATRIX(0, 13) = "PRICE (AS % OF " & Format(DATA_MATRIX(m, 1), "mmm d/yy") & " PRICE)"
TEMP_MATRIX(0, 16) = "(" & Format(MA1_PERIODS, "0") & " DAY) - (" & Format(MA2_PERIODS, "0") & _
                     " DAY) MOVING AVERAGE"
TEMP_MATRIX(0, 19) = "A TRENDLINE"
TEMP_MATRIX(0, 20) = "FIBONACCI FAN @" & Format(GOLDEN_RATIOS_VAL, "0.0%")
TEMP_MATRIX(0, 21) = "FIBONACCI FAN @" & Format(1 - GOLDEN_RATIOS_VAL, "0.0%")
GoSub TREND_LINES


'-----------------------------------------------------------------------------------
Select Case OUTPUT
'-----------------------------------------------------------------------------------
Case 0
'-----------------------------------------------------------------------------------
    ASSET_BOLLINGER_FIBONACCI_SIGNAL_FUNC = TEMP_MATRIX
'-----------------------------------------------------------------------------------
Case Else
'-----------------------------------------------------------------------------------
    If p = 0 Then: GoTo ERROR_LABEL
    MEAN_VAL = GTEMP_SUM / p
    SIGMA_VAL = 0
    For i = o To NROWS
        SIGMA_VAL = SIGMA_VAL + ((TEMP_MATRIX(i, 24) / TEMP_MATRIX(i - 1, 24) - 1) - MEAN_VAL) ^ 2
    Next i
    SIGMA_VAL = (SIGMA_VAL / p) ^ 0.5
    If OUTPUT = 1 Then
        ASSET_BOLLINGER_FIBONACCI_SIGNAL_FUNC = MEAN_VAL / SIGMA_VAL
    Else
        ASSET_BOLLINGER_FIBONACCI_SIGNAL_FUNC = Array(MEAN_VAL / SIGMA_VAL, MEAN_VAL, SIGMA_VAL)
    End If
'-----------------------------------------------------------------------------------
End Select
'-----------------------------------------------------------------------------------

Exit Function
'-----------------------------------------------------------------------------------
TREND_LINES:
'-----------------------------------------------------------------------------------
    If FIBONACCI_FLAG = True Then
        ATEMP_VAL = 2 ^ 52
        BTEMP_VAL = -2 ^ 52
    Else
        ATEMP_VAL = -2 ^ 52
        BTEMP_VAL = 2 ^ 52
    End If
    
    '-----------------------------------------------------------------------------------
    If FIBONACCI_FLAG = True Then
    '-----------------------------------------------------------------------------------
        For i = m + 1 To NROWS
            If TEMP_MATRIX(i, 13) < ATEMP_VAL Then
                ATEMP_VAL = TEMP_MATRIX(i, 13)
                ii = i
            End If
            If i > m + 2 Then
                TEMP_MATRIX(i, 19) = -1
                TEMP_MATRIX(i, 20) = -1
                TEMP_MATRIX(i, 21) = -1
            End If
        Next i
        For j = NROWS To ii + 1 Step -1
            If TEMP_MATRIX(j, 13) > BTEMP_VAL Then
                BTEMP_VAL = TEMP_MATRIX(j, 13)
                jj = j
            End If
        Next j
    '-----------------------------------------------------------------------------------
    Else
    '-----------------------------------------------------------------------------------
        For i = m + 1 To NROWS
            If TEMP_MATRIX(i, 13) > ATEMP_VAL Then
                ATEMP_VAL = TEMP_MATRIX(i, 13)
                ii = i
            End If
            If i > m + 2 Then
                TEMP_MATRIX(i, 19) = -1
                TEMP_MATRIX(i, 20) = -1
                TEMP_MATRIX(i, 21) = -1
            End If
        Next i
        For j = NROWS To ii + 1 Step -1
            If TEMP_MATRIX(j, 13) < BTEMP_VAL Then
                BTEMP_VAL = TEMP_MATRIX(j, 13)
                jj = j
            End If
        Next j
    '-----------------------------------------------------------------------------------
    End If
    '-----------------------------------------------------------------------------------
    SLOPE_VAL = (BTEMP_VAL - ATEMP_VAL) / (jj - ii)
    j = 0
    For i = ii To NROWS
        TEMP_MATRIX(i, 19) = ATEMP_VAL + SLOPE_VAL * j
        TEMP_MATRIX(i, 20) = ATEMP_VAL + (SLOPE_VAL * GOLDEN_RATIOS_VAL) * j
        TEMP_MATRIX(i, 21) = ATEMP_VAL + (SLOPE_VAL * (1 - GOLDEN_RATIOS_VAL)) * j
        j = j + 1
    Next i
'-----------------------------------------------------------------------------------
Return
'-----------------------------------------------------------------------------------

ERROR_LABEL:
ASSET_BOLLINGER_FIBONACCI_SIGNAL_FUNC = "--"
End Function

Function ASSET_BOLLINGER_FIBONACCI_SIGNAL_OPTIMIZER_FUNC(ByRef PARAM_RNG As Variant, _
ByRef CONST_RNG As Variant, _
ByRef TICKER_STR As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal END_DATE As Date, _
Optional ByVal REFERENCE_DATE As Date, _
Optional ByVal FIBONACCI_FLAG As Boolean = True, _
Optional ByVal INITIAL_CASH As Double = 1000)

On Error GoTo ERROR_LABEL

If IsArray(TICKER_STR) = True Then
    PUB_DATA_MATRIX = TICKER_STR
Else
    PUB_DATA_MATRIX = _
    YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, START_DATE, END_DATE, _
    "d", "DOHLCVA", False, True, True)
End If
PUB_REFERENCE_DATE = REFERENCE_DATE
PUB_FIBONACCI_FLAG = FIBONACCI_FLAG
PUB_INITIAL_CASH = INITIAL_CASH

'ASSET_BOLLINGER_FIBONACCI_SIGNAL_OPTIMIZER_FUNC = _
    NELDER_MEAD_OPTIMIZATION_FRAME_FUNC("ASSET_BOLLINGER_FIBONACCI_SIGNAL_OBJ_FUNC", _
    PARAM_RNG, CONST_RNG, False, 0, 10000, 0.000000000001)

ASSET_BOLLINGER_FIBONACCI_SIGNAL_OPTIMIZER_FUNC = _
    PIKAIA_OPTIMIZATION_FUNC("ASSET_BOLLINGER_FIBONACCI_SIGNAL_OBJ_FUNC", _
    CONST_RNG, False, , , , , , , , , , , , , , 0)

Exit Function
ERROR_LABEL:
ASSET_BOLLINGER_FIBONACCI_SIGNAL_OPTIMIZER_FUNC = Err.number
End Function

Function ASSET_BOLLINGER_FIBONACCI_SIGNAL_OBJ_FUNC(ByRef PARAM_VECTOR As Variant)

Dim THETA_VAL As Variant

On Error GoTo ERROR_LABEL

THETA_VAL = _
    ASSET_BOLLINGER_FIBONACCI_SIGNAL_FUNC(PUB_DATA_MATRIX, , , PUB_REFERENCE_DATE, _
    PARAM_VECTOR(1, 1), _
    PARAM_VECTOR(2, 1), _
    PARAM_VECTOR(3, 1), _
    PARAM_VECTOR(4, 1), _
    PARAM_VECTOR(5, 1), _
    PARAM_VECTOR(6, 1), _
    PUB_FIBONACCI_FLAG, _
    PUB_INITIAL_CASH, 1)
    
If IsNumeric(THETA_VAL) = False Then: GoTo ERROR_LABEL

ASSET_BOLLINGER_FIBONACCI_SIGNAL_OBJ_FUNC = THETA_VAL

Exit Function
ERROR_LABEL:
ASSET_BOLLINGER_FIBONACCI_SIGNAL_OBJ_FUNC = 1 / 1E+100
End Function
