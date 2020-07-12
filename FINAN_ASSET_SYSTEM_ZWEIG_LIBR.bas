Attribute VB_Name = "FINAN_ASSET_SYSTEM_ZWEIG_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'//////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////
Private Const PUB_EPSILON As Double = 2 ^ 52
Private PUB_INITIAL_EQUITY As Double
Private PUB_COUNT_BASIS As Double
Private PUB_DATA_MATRIX As Variant

Private PUB_INITIAL_INVESTMENT As Double
Private PUB_INITIAL_MONEY_MARKET As Double
Private PUB_MONEY_MARKET_RATE As Double

Private PUB_OBJ_FUNC As Integer
'//////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////


'************************************************************************************
'************************************************************************************
'FUNCTION      : ASSET_MARTIN_ZWEIG1_SYSTEM_FUNC
'DESCRIPTION   :
'LIBRARY       : FINAN_ASSET
'GROUP         : SIGNAL
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'http://www.financialwebring.org/gummystuff/zweig.htm
'************************************************************************************
'************************************************************************************

'-----------------------------------------------------------------------------------------------
'http://www.gummy-stuff.org/zweig.htm
'http://www.gummy-stuff.org/zweig-VLIC.htm
'-----------------------------------------------------------------------------------------------


Function ASSETS_MARTIN_ZWEIG1_SYSTEM_FUNC(ByVal TICKERS_RNG As Variant, _
ByVal START_DATE As Date, _
ByVal END_DATE As Date, _
Optional ByVal INITIAL_EQUITY_RNG As Double = 10000, _
Optional ByVal BUY_RULE_RNG As Double = 0.0255, _
Optional ByVal SELL_RULE_RNG As Double = 0.02675, _
Optional ByVal COUNT_BASIS_RNG As Long = 360)

Dim i As Long
Dim j As Long
Dim k As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim HEADINGS_STR As String
Dim TEMP_MATRIX As Variant
Dim TICKERS_VECTOR As Variant

Dim BUY_RULE_VECTOR As Variant
Dim SELL_RULE_VECTOR As Variant
Dim COUNT_BASIS_VECTOR As Variant
Dim INITIAL_EQUITY_VECTOR As Variant

On Error GoTo ERROR_LABEL

If IsArray(TICKERS_RNG) = True Then
    TICKERS_VECTOR = TICKERS_RNG
    If UBound(TICKERS_VECTOR, 1) = 1 Then
        TICKERS_VECTOR = MATRIX_TRANSPOSE_FUNC(TICKERS_VECTOR)
    End If
Else
    ReDim TICKERS_VECTOR(1 To 1, 1 To 1)
    TICKERS_VECTOR(1, 1) = TICKERS_RNG
End If
NROWS = UBound(TICKERS_VECTOR, 1)

If IsArray(INITIAL_EQUITY_RNG) = True Then
    INITIAL_EQUITY_VECTOR = INITIAL_EQUITY_RNG
    If UBound(INITIAL_EQUITY_VECTOR, 1) = 1 Then
        INITIAL_EQUITY_VECTOR = MATRIX_TRANSPOSE_FUNC(INITIAL_EQUITY_VECTOR)
    End If
Else
    ReDim INITIAL_EQUITY_VECTOR(1 To NROWS, 1 To 1)
    For i = 1 To NROWS
        INITIAL_EQUITY_VECTOR(i, 1) = INITIAL_EQUITY_RNG
    Next i
End If
If NROWS <> UBound(INITIAL_EQUITY_VECTOR, 1) Then: GoTo ERROR_LABEL

If IsArray(BUY_RULE_RNG) = True Then
    BUY_RULE_VECTOR = BUY_RULE_RNG
    If UBound(BUY_RULE_VECTOR, 1) = 1 Then
        BUY_RULE_VECTOR = MATRIX_TRANSPOSE_FUNC(BUY_RULE_VECTOR)
    End If
Else
    ReDim BUY_RULE_VECTOR(1 To NROWS, 1 To 1)
    For i = 1 To NROWS
        BUY_RULE_VECTOR(i, 1) = BUY_RULE_RNG
    Next i
End If
If NROWS <> UBound(BUY_RULE_VECTOR, 1) Then: GoTo ERROR_LABEL

If IsArray(SELL_RULE_RNG) = True Then
    SELL_RULE_VECTOR = SELL_RULE_RNG
    If UBound(SELL_RULE_VECTOR, 1) = 1 Then
        SELL_RULE_VECTOR = MATRIX_TRANSPOSE_FUNC(SELL_RULE_VECTOR)
    End If
Else
    ReDim SELL_RULE_VECTOR(1 To NROWS, 1 To 1)
    For i = 1 To NROWS
        SELL_RULE_VECTOR(i, 1) = SELL_RULE_RNG
    Next i
End If
If NROWS <> UBound(SELL_RULE_VECTOR, 1) Then: GoTo ERROR_LABEL

If IsArray(COUNT_BASIS_RNG) = True Then
    COUNT_BASIS_VECTOR = COUNT_BASIS_RNG
    If UBound(COUNT_BASIS_VECTOR, 1) = 1 Then
        COUNT_BASIS_VECTOR = MATRIX_TRANSPOSE_FUNC(COUNT_BASIS_VECTOR)
    End If
Else
    ReDim COUNT_BASIS_VECTOR(1 To NROWS, 1 To 1)
    For i = 1 To NROWS
        COUNT_BASIS_VECTOR(i, 1) = COUNT_BASIS_RNG
    Next i
End If
If NROWS <> UBound(COUNT_BASIS_VECTOR, 1) Then: GoTo ERROR_LABEL
GoSub REDIM_LINE
For i = 1 To NROWS
    TEMP_MATRIX(i, 1) = TICKERS_VECTOR(i, 1)
    TEMP_VECTOR = ASSET_MARTIN_ZWEIG1_SYSTEM_FUNC(TICKERS_VECTOR(i, 1), _
                  START_DATE, END_DATE, INITIAL_EQUITY_VECTOR(i, 1), BUY_RULE_VECTOR(i, 1), _
                  SELL_RULE_VECTOR(i, 1), , , , , COUNT_BASIS_VECTOR(i, 1), 1)
    If IsArray(TEMP_VECTOR) = False Then: GoTo 1983
    For k = 1 To 14
        TEMP_MATRIX(i, k + 1) = TEMP_VECTOR(k, 2)
    Next k
Next i

ASSETS_MARTIN_ZWEIG1_SYSTEM_FUNC = TEMP_MATRIX

Exit Function
'------------------------------------------------------------------------------------------
REDIM_LINE:
'------------------------------------------------------------------------------------------
    NCOLUMNS = 15
    ReDim TEMP_MATRIX(0 To NROWS, 1 To NCOLUMNS)
    HEADINGS_STR = "SYMBOL,START_DATE,END_DATE,DATA_POINTS,# OF TRADES,# OF TRADES / PERIOD,BUY RULE %,SELL RULE %,CURRENT SIGNAL,INITIAL EQUITY,BUY & HOLD BALANCE,BUY & CASH BALANCE,BUY & SELL BALANCE,2BUY & CASH BALANCE,2BUY & 2SELL BALANCE,"
    i = 1
    For k = 1 To NCOLUMNS
        j = InStr(i, HEADINGS_STR, ",")
        TEMP_MATRIX(0, k) = Mid(HEADINGS_STR, i, j - i)
        i = j + 1
    Next k
'------------------------------------------------------------------------------------------
Return
'------------------------------------------------------------------------------------------
ERROR_LABEL:
ASSETS_MARTIN_ZWEIG1_SYSTEM_FUNC = Err.number
End Function


Function ASSET_MARTIN_ZWEIG1_SYSTEM_FUNC(ByVal TICKER_STR As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal END_DATE As Date, _
Optional ByVal INITIAL_EQUITY As Double = 10000, _
Optional ByVal BUY_RULE As Double = 0.127675000074465, _
Optional ByVal SELL_RULE As Double = 0.082105000083579, _
Optional ByVal MIN_BUY_VAL As Double = 10 ^ -10, _
Optional ByVal MAX_BUY_VAL As Double = 0.5, _
Optional ByVal MIN_SELL_VAL As Double = 10 ^ -10, _
Optional ByVal MAX_SELL_VAL As Double = 0.5, _
Optional ByVal COUNT_BASIS As Long = 360, _
Optional ByVal OUTPUT As Integer = 0)

'Optimum % Rule will change as future data is entered
'Might be better to use the trend line for % Value to determine optimum
'Or might be better to use optimum but count on min in that area of graph
'This works better with more volitle indexies - there is a more
'distinctive peak

Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long
Dim NROWS As Long
Dim NCOLUMNS As Long
Dim TENOR_VAL As Double
Dim HEADINGS_STR As String
'Dim CONST_BOX As Variant
Dim PARAM_VECTOR As Variant
Dim TEMP_MATRIX As Variant
Dim TEMP_VECTOR As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

If IsArray(TICKER_STR) = False Then
    DATA_MATRIX = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, START_DATE, END_DATE, _
                  "w", "DOHLCVA", False, True, True)
Else
    DATA_MATRIX = TICKER_STR
End If

'---------------------------------------------------------------------------------------
If OUTPUT > 2 Then
    PUB_DATA_MATRIX = DATA_MATRIX
    PUB_INITIAL_EQUITY = INITIAL_EQUITY
    PUB_COUNT_BASIS = COUNT_BASIS
    Select Case OUTPUT
    Case 3 'BUY & CASH BALANCE
        PUB_OBJ_FUNC = 0
    Case 4 'BUY & SELL BALANCE
        PUB_OBJ_FUNC = 1
    Case 5 '2BUY & CASH BALANCE
        PUB_OBJ_FUNC = 2
    Case Else '2BUY & 2SELL BALANCE
        PUB_OBJ_FUNC = 3
    End Select
    GoSub OPTIMIZER_LINE
    ASSET_MARTIN_ZWEIG1_SYSTEM_FUNC = Array(PARAM_VECTOR(1, 1), PARAM_VECTOR(2, 1), ASSET_MARTIN_ZWEIG1_OBJ_FUNC(PARAM_VECTOR)) 'Buy Rule/Sell Rule/Balance
    Exit Function
'---------------------------------------------------------------------------------------
End If
'---------------------------------------------------------------------------------------
NROWS = UBound(DATA_MATRIX, 1)
GoSub REDIM_LINE
NCOLUMNS = UBound(DATA_MATRIX, 2)
'---------------------------------------------------------------------------------------
For i = 1 To NROWS - 1
    For j = 1 To NCOLUMNS
        TEMP_MATRIX(i, j + 1) = DATA_MATRIX(i, j)
    Next j
    TEMP_MATRIX(i, 7) = TEMP_MATRIX(i, 7) / 1000
    TEMP_MATRIX(i, 1) = DATA_MATRIX(i + 1, 1)
Next i
'---------------------------------------------------------------------------------------
i = NROWS
For j = 1 To NCOLUMNS: TEMP_MATRIX(i, j + 1) = DATA_MATRIX(i, j): Next j
TEMP_MATRIX(i, 7) = TEMP_MATRIX(i, 7) / 1000
TEMP_MATRIX(i, 1) = DATA_MATRIX(i, 1)
'---------------------------------------------------------------------------------------
l = 6 '--> Using Closing Prices '8
TEMP_MATRIX(1, 9) = ""
TEMP_MATRIX(1, 10) = TEMP_MATRIX(1, l)
TEMP_MATRIX(1, 11) = TEMP_MATRIX(1, l) / TEMP_MATRIX(1, 10) - 1
TEMP_MATRIX(1, 12) = "BUY"
TEMP_MATRIX(1, 13) = INITIAL_EQUITY
TEMP_MATRIX(1, 14) = INITIAL_EQUITY
TEMP_MATRIX(1, 15) = INITIAL_EQUITY
TEMP_MATRIX(1, 16) = INITIAL_EQUITY
TEMP_MATRIX(1, 17) = INITIAL_EQUITY
'---------------------------------------------------------------------------------------
For i = 2 To NROWS
        
    TEMP_MATRIX(i, 9) = TEMP_MATRIX(i, l) / TEMP_MATRIX(i - 1, l) - 1

    If TEMP_MATRIX(i - 1, 12) = "BUY" Then
        If TEMP_MATRIX(i - 1, 10) > TEMP_MATRIX(i, l) Then
            TEMP_MATRIX(i, 10) = TEMP_MATRIX(i - 1, 10)
        Else
            If TEMP_MATRIX(i - 1, l) > TEMP_MATRIX(i, l) Then
                TEMP_MATRIX(i, 10) = TEMP_MATRIX(i - 1, l)
            Else
                TEMP_MATRIX(i, 10) = TEMP_MATRIX(i, l)
            End If
        End If
    Else
        If TEMP_MATRIX(i - 1, 10) < TEMP_MATRIX(i, l) Then
            TEMP_MATRIX(i, 10) = TEMP_MATRIX(i - 1, 10)
        Else
            If TEMP_MATRIX(i - 1, l) < TEMP_MATRIX(i, l) Then
                TEMP_MATRIX(i, 10) = TEMP_MATRIX(i - 1, l)
            Else
                TEMP_MATRIX(i, 10) = TEMP_MATRIX(i, l)
            End If
        End If
    End If

    TEMP_MATRIX(i, 11) = TEMP_MATRIX(i, l) / TEMP_MATRIX(i, 10) - 1
    
    If TEMP_MATRIX(i, 11) <= -SELL_RULE Then
        TEMP_MATRIX(i, 12) = "SELL"
    Else
        If TEMP_MATRIX(i, 11) >= BUY_RULE Then
            TEMP_MATRIX(i, 12) = "BUY"
        Else
            TEMP_MATRIX(i, 12) = TEMP_MATRIX(i - 1, 12)
        End If
    End If
        
    TEMP_MATRIX(i, 13) = TEMP_MATRIX(i - 1, 13) * TEMP_MATRIX(i, l) / TEMP_MATRIX(i - 1, l)
        
    If TEMP_MATRIX(i - 1, 12) = "BUY" Then
        If TEMP_MATRIX(i - 1, 12) <> TEMP_MATRIX(i - 2, 12) Then
            TEMP_MATRIX(i, 14) = TEMP_MATRIX(i - 1, 14) * TEMP_MATRIX(i, l) / TEMP_MATRIX(i, 3)
            TEMP_MATRIX(i, 15) = TEMP_MATRIX(i - 1, 15) * TEMP_MATRIX(i, l) / TEMP_MATRIX(i, 3)
            TEMP_MATRIX(i, 16) = TEMP_MATRIX(i - 1, 16) * (1 + ((TEMP_MATRIX(i, l) / _
                                 TEMP_MATRIX(i, 3)) - 1) * 2)
            TEMP_MATRIX(i, 17) = TEMP_MATRIX(i - 1, 17) * (1 + ((TEMP_MATRIX(i, l) / _
                                 TEMP_MATRIX(i, 3) - 1) * 2))
        Else
            TEMP_MATRIX(i, 14) = TEMP_MATRIX(i - 1, 14) * TEMP_MATRIX(i, l) / TEMP_MATRIX(i - 1, l)
            TEMP_MATRIX(i, 15) = TEMP_MATRIX(i - 1, 15) * TEMP_MATRIX(i, l) / TEMP_MATRIX(i - 1, l)
            TEMP_MATRIX(i, 16) = TEMP_MATRIX(i - 1, 16) * (1 + (TEMP_MATRIX(i, 9) * 2))
            TEMP_MATRIX(i, 17) = TEMP_MATRIX(i - 1, 17) * (1 + ((TEMP_MATRIX(i, l) / _
                                 TEMP_MATRIX(i - 1, l) - 1) * 2))
        End If
    Else
        TEMP_MATRIX(i, 14) = TEMP_MATRIX(i - 1, 14)
        If TEMP_MATRIX(i - 1, 12) <> TEMP_MATRIX(i - 2, 12) Then
           TEMP_MATRIX(i, 15) = TEMP_MATRIX(i - 1, 15) * ((1 + (-1 * (TEMP_MATRIX(i, 3) / _
                                TEMP_MATRIX(i - 1, l) - 1))))
           TEMP_MATRIX(i, 17) = TEMP_MATRIX(i - 1, 17) * ((1 + (-1 * (TEMP_MATRIX(i, 3) / _
                                TEMP_MATRIX(i - 1, l) - 1))))
        Else
            TEMP_MATRIX(i, 15) = TEMP_MATRIX(i - 1, 15) * ((1 + (-1 * TEMP_MATRIX(i, 9))))
            TEMP_MATRIX(i, 17) = TEMP_MATRIX(i - 1, 17) * ((1 + (-1 * TEMP_MATRIX(i, 9) * 2)))
        End If
        TEMP_MATRIX(i, 16) = TEMP_MATRIX(i - 1, 16)
    End If
  
Next i
'---------------------------------------------------------------------------------------
k = 0
i = NROWS
TEMP_MATRIX(i, 18) = Year(TEMP_MATRIX(i, 1))
GoSub TRADE_LINE

For i = NROWS - 1 To 1 Step -1
    If TEMP_MATRIX(i + 1, 1) = "" Then
        TEMP_MATRIX(i, 18) = Year(TEMP_MATRIX(i, 1))
    Else
        If Year(TEMP_MATRIX(i + 1, 1)) <> Year(TEMP_MATRIX(i, 1)) Then
            TEMP_MATRIX(i, 18) = Year(TEMP_MATRIX(i, 1))
        Else
            TEMP_MATRIX(i, 18) = ""
        End If
    End If
    GoSub TRADE_LINE
Next i
'---------------------------------------------------------------------------------------
'------------------------------------------------------------------------------
Select Case OUTPUT
'------------------------------------------------------------------------------
Case 0
'------------------------------------------------------------------------------
    ASSET_MARTIN_ZWEIG1_SYSTEM_FUNC = TEMP_MATRIX
'------------------------------------------------------------------------------
Case 1
'------------------------------------------------------------------------------
    ReDim TEMP_VECTOR(1 To 14, 1 To 2)

    TEMP_VECTOR(1, 1) = "START_DATE"
    TEMP_VECTOR(1, 2) = TEMP_MATRIX(1, 2)
            
    TEMP_VECTOR(2, 1) = "END_DATE"
    TEMP_VECTOR(2, 2) = TEMP_MATRIX(NROWS, 1)
            
    TEMP_VECTOR(3, 1) = "DATA_POINTS"
    TEMP_VECTOR(3, 2) = NROWS
            
    TEMP_VECTOR(4, 1) = "# OF TRADES"
    TEMP_VECTOR(4, 2) = k
            
    TEMP_VECTOR(5, 1) = "# OF TRADES / PERIOD"
    TENOR_VAL = DateDiff("d", TEMP_VECTOR(1, 2), TEMP_VECTOR(2, 2)) / COUNT_BASIS
    TEMP_VECTOR(5, 2) = TEMP_VECTOR(4, 2) / TENOR_VAL
            
    TEMP_VECTOR(6, 1) = "BUY RULE %"
    TEMP_VECTOR(6, 2) = BUY_RULE
            
    TEMP_VECTOR(7, 1) = "SELL RULE %"
    TEMP_VECTOR(7, 2) = SELL_RULE
            
    TEMP_VECTOR(8, 1) = "CURRENT SIGNAL"
    TEMP_VECTOR(8, 2) = TEMP_MATRIX(NROWS, 12)
            
    TEMP_VECTOR(9, 1) = "INITIAL EQUITY"
    TEMP_VECTOR(9, 2) = INITIAL_EQUITY
            
    TEMP_VECTOR(10, 1) = "BUY & HOLD BALANCE"
    TEMP_VECTOR(10, 2) = TEMP_MATRIX(NROWS, 13)
            
    TEMP_VECTOR(11, 1) = "BUY & CASH BALANCE"
    TEMP_VECTOR(11, 2) = TEMP_MATRIX(NROWS, 14)
            
    TEMP_VECTOR(12, 1) = "BUY & SELL BALANCE"
    TEMP_VECTOR(12, 2) = TEMP_MATRIX(NROWS, 15)
            
    TEMP_VECTOR(13, 1) = "2BUY & CASH BALANCE"
    TEMP_VECTOR(13, 2) = TEMP_MATRIX(NROWS, 16)
            
    TEMP_VECTOR(14, 1) = "2BUY & 2SELL BALANCE"
    TEMP_VECTOR(14, 2) = TEMP_MATRIX(NROWS, 17)
        
    ASSET_MARTIN_ZWEIG1_SYSTEM_FUNC = TEMP_VECTOR
'------------------------------------------------------------------------------
Case 2 'End of Year Cumulative Summary
'------------------------------------------------------------------------------
    ASSET_MARTIN_ZWEIG1_SYSTEM_FUNC = MATRIX_TRIM_FUNC((ARRAY_GET_VECTOR_FUNC(TEMP_MATRIX, 18, 23, 0, NROWS)), 1, "")
'------------------------------------------------------------------------------
End Select
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
Exit Function
'------------------------------------------------------------------------------------------
REDIM_LINE:
'------------------------------------------------------------------------------------------
    NCOLUMNS = 24
    ReDim TEMP_MATRIX(0 To NROWS, 1 To NCOLUMNS)
    HEADINGS_STR = "END PERIOD,START PERIOD,OPEN,HIGH,LOW,CLOSE,VOLUME,ADJ.CLOSE,RETURN,MIN/MAX,CLOSE VS MIN/MAX,SIGNAL,BUY AND HOLD,BUY AND CASH,BUY AND SELL,2BUY AND CASH,2BUY AND 2SELL,END OF YEAR,BUY AND HOLD,BUY AND CASH,BUY AND SELL,2BUY AND CASH,2BUY AND 2SELL,NO TRADES,"
    i = 1
    For k = 1 To NCOLUMNS
        j = InStr(i, HEADINGS_STR, ",")
        TEMP_MATRIX(0, k) = Mid(HEADINGS_STR, i, j - i)
        i = j + 1
    Next k
'------------------------------------------------------------------------------------------
Return
'------------------------------------------------------------------------------
OPTIMIZER_LINE:
'------------------------------------------------------------------------------
    'ReDim PARAM_VECTOR(1 To 2, 1 To 1)
    'PARAM_VECTOR(1, 1) = BUY_RULE
    'PARAM_VECTOR(2, 1) = SELL_RULE
    
    If MIN_BUY_VAL < 0 Or MIN_BUY_VAL > 1 Then: MIN_BUY_VAL = 10 ^ -10
    If MIN_SELL_VAL < 0 Or MIN_SELL_VAL > 1 Then: MIN_SELL_VAL = 10 ^ -10
    
    If MAX_BUY_VAL < 0 Or MAX_BUY_VAL > 1 Then: MAX_BUY_VAL = 1
    If MAX_SELL_VAL < 0 Or MAX_SELL_VAL > 1 Then: MAX_SELL_VAL = 1
    
    ReDim CONST_BOX(1 To 2, 1 To 2)
    CONST_BOX(1, 1) = MIN_BUY_VAL
    CONST_BOX(2, 1) = MAX_BUY_VAL 'buy-max
        
    CONST_BOX(1, 2) = MIN_SELL_VAL
    CONST_BOX(2, 2) = MAX_SELL_VAL 'sell-max
        
    PARAM_VECTOR = PIKAIA_OPTIMIZATION_FUNC("ASSET_MARTIN_ZWEIG1_OBJ_FUNC", _
                   CONST_BOX, False, , , , , , , , , , , , , , 0)

'   PARAM_VECTOR = NELDER_MEAD_OPTIMIZATION_FRAME_FUNC("ASSET_MARTIN_ZWEIG1_OBJ_FUNC", _
                   PARAM_VECTOR, CONST_BOX, False, 0, 1000, 10 ^ -15)
'   PARAM_VECTOR = NELDER_MEAD_OPTIMIZATION_FRAME_FUNC("ASSET_MARTIN_ZWEIG1_OBJ_FUNC", _
                   PARAM_VECTOR, , False, 0, 1000, 10 ^ -15)
'------------------------------------------------------------------------------
Return
'------------------------------------------------------------------------------------------
TRADE_LINE:
'------------------------------------------------------------------------------------------
    If TEMP_MATRIX(i, 18) <> "" Then
        For j = 19 To 23: TEMP_MATRIX(i, j) = TEMP_MATRIX(i, j - 6): Next j
    Else
        For j = 19 To 23: TEMP_MATRIX(i, j) = "": Next j
    End If
        
    If TEMP_MATRIX(i, 2) <> "" Then
        If TEMP_MATRIX(i, 12) <> TEMP_MATRIX(i - 1, 12) Then
            TEMP_MATRIX(i, 24) = 1
            k = k + 1
        Else
            TEMP_MATRIX(i, 24) = 0
        End If
    Else
        TEMP_MATRIX(i, 24) = ""
    End If
'------------------------------------------------------------------------------------------
Return
'------------------------------------------------------------------------------------------
ERROR_LABEL:
ASSET_MARTIN_ZWEIG1_SYSTEM_FUNC = PUB_EPSILON
End Function

Function ASSET_MARTIN_ZWEIG1_SENSTIVITY_FUNC(ByVal TICKER_STR As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal END_DATE As Date, _
Optional ByVal INITIAL_EQUITY As Double = 10000, _
Optional ByVal MIN_BUY_VAL As Double = 0.01, _
Optional ByVal MAX_BUY_VAL As Double = 0.2, _
Optional ByVal DELTA_BUY_RULE As Double = 0.01, _
Optional ByVal MIN_SELL_VAL As Double = 0.01, _
Optional ByVal MAX_SELL_VAL As Double = 0.2, _
Optional ByVal DELTA_SELL_RULE As Double = 0.01, _
Optional ByVal COUNT_BASIS As Long = 360)

Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim NSIZE As Long
Dim NROWS As Long
Dim NCOLUMNS As Long

Dim BUY_RULE As Double
Dim SELL_RULE As Double

Dim HEADINGS_STR As String
Dim TEMP_VECTOR As Variant
Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

If IsArray(TICKER_STR) = False Then
    DATA_MATRIX = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, START_DATE, END_DATE, _
                  "w", "DOHLCVA", False, True, True)
Else
    DATA_MATRIX = TICKER_STR
End If

NROWS = (MAX_BUY_VAL - MIN_BUY_VAL) / DELTA_BUY_RULE + 1
NCOLUMNS = (MAX_SELL_VAL - MIN_SELL_VAL) / DELTA_SELL_RULE + 1
NSIZE = NROWS * NCOLUMNS
GoSub REDIM_LINE
BUY_RULE = MIN_BUY_VAL
l = 1
For i = 1 To NROWS
    SELL_RULE = MIN_SELL_VAL
    For j = 1 To NCOLUMNS
        TEMP_VECTOR = ASSET_MARTIN_ZWEIG1_SYSTEM_FUNC(DATA_MATRIX, , , INITIAL_EQUITY, BUY_RULE, SELL_RULE, , , , , COUNT_BASIS, 1)
        If IsArray(TEMP_VECTOR) = False Then: GoTo 1983
        For k = 1 To 14: TEMP_MATRIX(l, k) = TEMP_VECTOR(k, 2): Next k
        SELL_RULE = SELL_RULE + DELTA_SELL_RULE
1983:
        l = l + 1
    Next j
    BUY_RULE = BUY_RULE + DELTA_BUY_RULE
Next i

ASSET_MARTIN_ZWEIG1_SENSTIVITY_FUNC = TEMP_MATRIX

Exit Function
'------------------------------------------------------------------------------------------
REDIM_LINE:
'------------------------------------------------------------------------------------------
    NCOLUMNS = 14
    ReDim TEMP_MATRIX(0 To NSIZE, 1 To NCOLUMNS)
    HEADINGS_STR = "START_DATE,END_DATE,DATA_POINTS,# OF TRADES,# OF TRADES / PERIOD,BUY RULE %,SELL RULE %,CURRENT SIGNAL,INITIAL EQUITY,BUY & HOLD BALANCE,BUY & CASH BALANCE,BUY & SELL BALANCE,2BUY & CASH BALANCE,2BUY & 2SELL BALANCE,"
    i = 1
    For k = 1 To NCOLUMNS
        j = InStr(i, HEADINGS_STR, ",")
        TEMP_MATRIX(0, k) = Mid(HEADINGS_STR, i, j - i)
        i = j + 1
    Next k
'------------------------------------------------------------------------------------------
Return
'------------------------------------------------------------------------------------------
ERROR_LABEL:
ASSET_MARTIN_ZWEIG1_SENSTIVITY_FUNC = Err.number
End Function

'Optimizer for Single
Private Function ASSET_MARTIN_ZWEIG1_OBJ_FUNC(ByRef PARAM_VECTOR As Variant)

Dim i As Long
Dim TEMP_VECTOR As Variant

On Error GoTo ERROR_LABEL
    
Select Case PUB_OBJ_FUNC
Case 0 'BUY & CASH BALANCE
    i = 11
Case 1 'BUY & SELL BALANCE
    i = 12
Case 2 '2BUY & CASH BALANCE
    i = 13
Case 3 '2BUY & 2SELL BALANCE
    i = 14
End Select

TEMP_VECTOR = ASSET_MARTIN_ZWEIG1_SYSTEM_FUNC(PUB_DATA_MATRIX, , , PUB_INITIAL_EQUITY, _
PARAM_VECTOR(1, 1), PARAM_VECTOR(2, 1), , , , , PUB_COUNT_BASIS, 1)
If IsArray(TEMP_VECTOR) = False Then: GoTo ERROR_LABEL

ASSET_MARTIN_ZWEIG1_OBJ_FUNC = TEMP_VECTOR(i, 2)

Exit Function
ERROR_LABEL:
ASSET_MARTIN_ZWEIG1_OBJ_FUNC = PUB_EPSILON
End Function

Function ASSET_MARTIN_ZWEIG2_SYSTEM_FUNC(ByVal TICKER_STR As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal END_DATE As Date, _
Optional ByVal VERSION As Integer = 3, _
Optional ByVal INITIAL_INVESTMENT As Double = 10000, _
Optional ByVal INITIAL_MONEY_MARKET As Double = 10000, _
Optional ByVal MONEY_MARKET_RATE As Double = 0.02, _
Optional ByVal BUY_VAL As Double = 0.011245000097751, _
Optional ByVal SELL_VAL As Double = 0.015215000096957, _
Optional ByVal MIN_BUY_VAL As Double = 10 ^ -10, _
Optional ByVal MAX_BUY_VAL As Double = 0.5, _
Optional ByVal MIN_SELL_VAL As Double = 10 ^ -10, _
Optional ByVal MAX_SELL_VAL As Double = 0.5, _
Optional ByVal COUNT_BASIS As Double = 52, _
Optional ByVal FREQUENCY_STR As String = "W", _
Optional ByVal OUTPUT As Integer = 0)

'BUY_VAL: down % from Hi
'SELL_VAL: up % from Lo

Dim i As Long
Dim j As Long
Dim k As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim MEAN_VAL As Double
Dim VOLAT_VAL As Double
Dim FACTOR_VAL As Double

Const BUY_STR As String = "BUY"
Const SELL_STR As String = "SELL"
Dim HEADINGS_STR As String

Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

If IsArray(TICKER_STR) = False Then
    DATA_MATRIX = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, START_DATE, END_DATE, _
                  FREQUENCY_STR, "DOHLCVA", False, True, True)
Else
    DATA_MATRIX = TICKER_STR
End If
NROWS = UBound(DATA_MATRIX, 1)

If OUTPUT > 2 Then
    PUB_DATA_MATRIX = DATA_MATRIX
    PUB_INITIAL_INVESTMENT = INITIAL_INVESTMENT
    PUB_INITIAL_MONEY_MARKET = INITIAL_MONEY_MARKET
    PUB_MONEY_MARKET_RATE = MONEY_MARKET_RATE
    PUB_OBJ_FUNC = VERSION
    PUB_COUNT_BASIS = COUNT_BASIS

    If MIN_BUY_VAL < 0 Or MIN_BUY_VAL > 1 Then: MIN_BUY_VAL = 10 ^ -10
    If MIN_SELL_VAL < 0 Or MIN_SELL_VAL > 1 Then: MIN_SELL_VAL = 10 ^ -10
    
    If MAX_BUY_VAL < 0 Or MAX_BUY_VAL > 1 Then: MAX_BUY_VAL = 1
    If MAX_SELL_VAL < 0 Or MAX_SELL_VAL > 1 Then: MAX_SELL_VAL = 1
    
    ReDim CONST_BOX(1 To 2, 1 To 2)
    CONST_BOX(1, 1) = MIN_BUY_VAL
    CONST_BOX(2, 1) = MAX_BUY_VAL
        
    CONST_BOX(1, 2) = MIN_SELL_VAL
    CONST_BOX(2, 2) = MAX_SELL_VAL
    
    ASSET_MARTIN_ZWEIG2_SYSTEM_FUNC = PIKAIA_OPTIMIZATION_FUNC("ASSET_MARTIN_ZWEIG2_OBJ_FUNC", _
    CONST_BOX, False, , , , , , , , , , , , , , 0)
    Exit Function
End If

FACTOR_VAL = (1 + MONEY_MARKET_RATE) ^ (1 / COUNT_BASIS)
GoSub REDIM_LINE
TEMP_MATRIX(0, 14) = INITIAL_INVESTMENT
TEMP_MATRIX(0, 15) = INITIAL_MONEY_MARKET

i = 1
For j = 1 To 7: TEMP_MATRIX(i, j) = DATA_MATRIX(i, j): Next j
TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 6) / 1000
TEMP_MATRIX(i, 8) = TEMP_MATRIX(i, 5) / TEMP_MATRIX(i, 2)
GoSub SIGNAL_LINE
GoSub BALANCE_LINE
TEMP_MATRIX(i, 13) = TEMP_MATRIX(i, 14) + TEMP_MATRIX(i, 15)
MEAN_VAL = TEMP_MATRIX(i, 13) / (INITIAL_MONEY_MARKET + INITIAL_INVESTMENT) - 1
TEMP_MATRIX(i, 16) = TEMP_MATRIX(i, 8) * (INITIAL_MONEY_MARKET + INITIAL_INVESTMENT)

For i = 2 To NROWS
    For j = 1 To 7: TEMP_MATRIX(i, j) = DATA_MATRIX(i, j): Next j
    TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 6) / 1000
    TEMP_MATRIX(i, 8) = TEMP_MATRIX(i, 7) / TEMP_MATRIX(i - 1, 7)
    GoSub SIGNAL_LINE
    GoSub BALANCE_LINE
    TEMP_MATRIX(i, 13) = TEMP_MATRIX(i, 14) + TEMP_MATRIX(i, 15)
    MEAN_VAL = MEAN_VAL + TEMP_MATRIX(i, 13) / TEMP_MATRIX(i - 1, 13) - 1
    TEMP_MATRIX(i, 16) = TEMP_MATRIX(i, 8) * TEMP_MATRIX(i - 1, 16)
Next i
TEMP_MATRIX(0, 14) = "SYSTEM BALANCE"
TEMP_MATRIX(0, 15) = "CASH BALANCE: " & Format(TEMP_MATRIX(0, 15), "$#,#00.0")

If OUTPUT = 0 Then
    ASSET_MARTIN_ZWEIG2_SYSTEM_FUNC = TEMP_MATRIX
Else
    MEAN_VAL = MEAN_VAL / NROWS
    i = 1
    VOLAT_VAL = ((TEMP_MATRIX(i, 13) / (INITIAL_MONEY_MARKET + INITIAL_INVESTMENT) - 1) - MEAN_VAL) ^ 2
    For i = 2 To NROWS
        VOLAT_VAL = VOLAT_VAL + ((TEMP_MATRIX(i, 13) / TEMP_MATRIX(i - 1, 13) - 1) - MEAN_VAL) ^ 2
    Next i
    VOLAT_VAL = (VOLAT_VAL / NROWS) ^ 0.5
    
    If OUTPUT = 1 Then
        ASSET_MARTIN_ZWEIG2_SYSTEM_FUNC = Array(MEAN_VAL / VOLAT_VAL, MEAN_VAL, VOLAT_VAL)
    Else
        ASSET_MARTIN_ZWEIG2_SYSTEM_FUNC = (MEAN_VAL / VOLAT_VAL)
    End If
End If

'------------------------------------------------------------------------------------------
Exit Function
'------------------------------------------------------------------------------------------
REDIM_LINE:
'------------------------------------------------------------------------------------------
    NCOLUMNS = 16
    ReDim TEMP_MATRIX(0 To NROWS, 1 To NCOLUMNS)
    HEADINGS_STR = "DATE,OPEN,HIGH,LOW,CLOSE,VOLUME,ADJ.CLOSE,GAIN,SIGNAL,BUY @,SELL @,TRADES,PORTFOLIO BALANCE,SYSTEM BALANCE,CASH BALANCE,BUY HOLD BALANCE," 'BUY HOLD GAIN,"
    i = 1
    For k = 1 To NCOLUMNS
        j = InStr(i, HEADINGS_STR, ",")
        TEMP_MATRIX(0, k) = Mid(HEADINGS_STR, i, j - i)
        i = j + 1
    Next k
'------------------------------------------------------------------------------------------
Return
'------------------------------------------------------------------------------------------
BALANCE_LINE:
'------------------------------------------------------------------------------------------
    If TEMP_MATRIX(i, 9) = "" Then
        TEMP_MATRIX(i, 14) = TEMP_MATRIX(i - 1, 14) * TEMP_MATRIX(i, 8)
    Else
        If TEMP_MATRIX(i, 9) = BUY_STR Then
            TEMP_MATRIX(i, 14) = TEMP_MATRIX(i - 1, 14) + TEMP_MATRIX(i - 1, 15) * FACTOR_VAL
        Else
            If TEMP_MATRIX(i, 9) = SELL_STR Then
                TEMP_MATRIX(i, 14) = 0
            Else
                TEMP_MATRIX(i, 14) = ""
            End If
        End If
    End If
    '--------------------------------------------------------------------------------------
    If TEMP_MATRIX(i, 9) = "" Then
        TEMP_MATRIX(i, 15) = TEMP_MATRIX(i - 1, 15) * FACTOR_VAL
    Else
        If TEMP_MATRIX(i, 9) = BUY_STR Then
            TEMP_MATRIX(i, 15) = 0
        Else
            If TEMP_MATRIX(i, 9) = SELL_STR Then
                TEMP_MATRIX(i, 15) = TEMP_MATRIX(i - 1, 15) + TEMP_MATRIX(i - 1, 14) * TEMP_MATRIX(i, 8)
            Else
                TEMP_MATRIX(i, 15) = ""
            End If
        End If
    End If
    '--------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------
Return
'------------------------------------------------------------------------------------------
SIGNAL_LINE:
'------------------------------------------------------------------------------------------
    Select Case VERSION
    '--------------------------------------------------------------------------------------
    Case 0
    '--------------------------------------------------------------------------------------
        If (TEMP_MATRIX(i - 1, 15) > 0 And TEMP_MATRIX(i, 5) > (1 + SELL_VAL) * TEMP_MATRIX(i, 4)) Then
            TEMP_MATRIX(i, 9) = BUY_STR
        Else
            If (TEMP_MATRIX(i - 1, 14) > 0 And TEMP_MATRIX(i, 5) < (1 - BUY_VAL) * TEMP_MATRIX(i, 3)) Then
                TEMP_MATRIX(i, 9) = SELL_STR
            Else
                TEMP_MATRIX(i, 9) = ""
            End If
        End If
    '--------------------------------------------------------------------------------------
    Case 1
    '--------------------------------------------------------------------------------------
        If (TEMP_MATRIX(i - 1, 15) > 0 And TEMP_MATRIX(i, 5) < (1 + SELL_VAL) * TEMP_MATRIX(i, 4)) Then
            TEMP_MATRIX(i, 9) = BUY_STR
        Else
            If (TEMP_MATRIX(i - 1, 14) > 0 And TEMP_MATRIX(i, 5) > (1 - BUY_VAL) * TEMP_MATRIX(i, 3)) Then
                TEMP_MATRIX(i, 9) = SELL_STR
            Else
                TEMP_MATRIX(i, 9) = ""
            End If
        End If
    '--------------------------------------------------------------------------------------
    Case 2
    '--------------------------------------------------------------------------------------
        If (TEMP_MATRIX(i - 1, 15) > 0 And TEMP_MATRIX(i, 5) < (1 - BUY_VAL) * TEMP_MATRIX(i, 3)) Then
            TEMP_MATRIX(i, 9) = BUY_STR
        Else
            If (TEMP_MATRIX(i - 1, 14) > 0 And TEMP_MATRIX(i, 5) > (1 + SELL_VAL) * TEMP_MATRIX(i, 4)) Then
                TEMP_MATRIX(i, 9) = SELL_STR
            Else
                TEMP_MATRIX(i, 9) = ""
            End If
        End If
    '--------------------------------------------------------------------------------------
    Case Else
    '--------------------------------------------------------------------------------------
        If (TEMP_MATRIX(i - 1, 15) > 0 And TEMP_MATRIX(i, 5) > (1 - BUY_VAL) * TEMP_MATRIX(i, 3)) Then
            TEMP_MATRIX(i, 9) = BUY_STR
        Else
            If (TEMP_MATRIX(i - 1, 14) > 0 And TEMP_MATRIX(i, 5) < (1 + SELL_VAL) * TEMP_MATRIX(i, 4)) Then
                TEMP_MATRIX(i, 9) = SELL_STR
            Else
                TEMP_MATRIX(i, 9) = ""
            End If
        End If
    '--------------------------------------------------------------------------------------
    End Select
    '--------------------------------------------------------------------------------------
    If TEMP_MATRIX(i, 9) = BUY_STR Then
        TEMP_MATRIX(i, 10) = TEMP_MATRIX(i, 7)
        TEMP_MATRIX(i, 11) = 0
        TEMP_MATRIX(i, 12) = 1
    ElseIf TEMP_MATRIX(i, 9) = SELL_STR Then
        TEMP_MATRIX(i, 10) = 0
        TEMP_MATRIX(i, 11) = TEMP_MATRIX(i, 7)
        TEMP_MATRIX(i, 12) = 1
    Else
        TEMP_MATRIX(i, 10) = 0
        TEMP_MATRIX(i, 11) = 0
        TEMP_MATRIX(i, 12) = 0
    End If
'------------------------------------------------------------------------------------------
Return
'------------------------------------------------------------------------------------------
ERROR_LABEL:
'------------------------------------------------------------------------------------------
ASSET_MARTIN_ZWEIG2_SYSTEM_FUNC = PUB_EPSILON
'------------------------------------------------------------------------------------------
End Function

'Optimizer for Single
Private Function ASSET_MARTIN_ZWEIG2_OBJ_FUNC(ByRef PARAM_VECTOR As Variant)

Dim i As Long
Dim Y_VAL As Double

On Error GoTo ERROR_LABEL

Select Case PUB_OBJ_FUNC
Case 0
    i = 0
Case 1
    i = 1
Case 2
    i = 2
Case Else
    i = 3
End Select

Y_VAL = ASSET_MARTIN_ZWEIG2_SYSTEM_FUNC(PUB_DATA_MATRIX, , , i, PUB_INITIAL_INVESTMENT, _
        PUB_INITIAL_MONEY_MARKET, PUB_MONEY_MARKET_RATE, PARAM_VECTOR(1, 1), _
        PARAM_VECTOR(2, 1), , , , , PUB_COUNT_BASIS, 2)

If Y_VAL <> PUB_EPSILON Then
    ASSET_MARTIN_ZWEIG2_OBJ_FUNC = Y_VAL
Else
    GoTo ERROR_LABEL
End If

Exit Function
ERROR_LABEL:
ASSET_MARTIN_ZWEIG2_OBJ_FUNC = 1 / PUB_EPSILON
End Function
