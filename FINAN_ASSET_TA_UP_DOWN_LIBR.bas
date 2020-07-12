Attribute VB_Name = "FINAN_ASSET_TA_UP_DOWN_LIBR"

'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
Private PUB_VERSION As Integer
Private PUB_DATA_MATRIX As Variant
'------------------------------------------------------------------------------------

Function ASSETS_OPEN_PRICE_UP_DOWN_FUNC(ByRef TICKERS_RNG As Variant, _
ByVal START_DATE As Date, _
ByVal END_DATE As Date, _
Optional ByRef PERCENT_VAL_RNG As Variant)

Dim i As Long
Dim j As Long
Dim k As Long

Dim ii As Long
Dim jj As Long
Dim kk As Long

Dim NROWS As Long
Dim TICKER_STR As String
Dim PERCENT_VAL As Variant

Dim TEMP_ARR As Variant
Dim TEMP_MATRIX As Variant
Dim TICKERS_VECTOR As Variant
Dim PERCENT_VAL_VECTOR As Variant

On Error GoTo ERROR_LABEL

TICKERS_VECTOR = TICKERS_RNG
If UBound(TICKERS_VECTOR, 1) = 1 Then
    TICKERS_VECTOR = MATRIX_TRANSPOSE_FUNC(TICKERS_VECTOR)
End If
NROWS = UBound(TICKERS_VECTOR, 1)

If IsArray(PERCENT_VAL_RNG) = True Then
    PERCENT_VAL_VECTOR = PERCENT_VAL_RNG
    If UBound(PERCENT_VAL_VECTOR, 1) = 1 Then
        PERCENT_VAL_VECTOR = MATRIX_TRANSPOSE_FUNC(PERCENT_VAL_VECTOR)
    End If
Else
    ReDim PERCENT_VAL_VECTOR(1 To NROWS, 1 To 1)
    For i = 1 To NROWS: PERCENT_VAL_VECTOR(i, 1) = "": Next i
End If

ReDim TEMP_MATRIX(0 To NROWS, 1 To 17)
TEMP_MATRIX(0, 1) = "TICKER"
TEMP_MATRIX(0, 2) = "HIGHER: PERCENT_VAL"
TEMP_MATRIX(0, 3) = "HIGHER: AVERAGE"
TEMP_MATRIX(0, 4) = "HIGHER: VOLATILITY"
TEMP_MATRIX(0, 5) = "HIGHER: CAGR"
TEMP_MATRIX(0, 6) = "HIGHER: %C>O"
TEMP_MATRIX(0, 7) = "HIGHER: AVERAGE(C>O)"
TEMP_MATRIX(0, 8) = "HIGHER: %C<O"
TEMP_MATRIX(0, 9) = "HIGHER: AVERAGE(C<O)"

TEMP_MATRIX(0, 10) = "LOWER: PERCENT_VAL"
TEMP_MATRIX(0, 11) = "LOWER: AVERAGE"
TEMP_MATRIX(0, 12) = "LOWER: VOLATILITY"
TEMP_MATRIX(0, 13) = "LOWER: CAGR"
TEMP_MATRIX(0, 14) = "LOWER: %C>O"
TEMP_MATRIX(0, 15) = "LOWER: AVERAGE(C>O)"
TEMP_MATRIX(0, 16) = "LOWER: %C<O"
TEMP_MATRIX(0, 17) = "LOWER: AVERAGE(C<O)"

For i = 1 To NROWS
    TICKER_STR = TICKERS_VECTOR(i, 1)
    TEMP_MATRIX(i, 1) = TICKER_STR
    PERCENT_VAL = PERCENT_VAL_VECTOR(i, 1)
    ii = 0: jj = 2: kk = 9
    GoSub SENSITIVITY_LINE
    ii = 1: jj = 10: kk = 17
    GoSub SENSITIVITY_LINE
Next i
ASSETS_OPEN_PRICE_UP_DOWN_FUNC = TEMP_MATRIX

'------------------------------------------------------------------------------------------------------
Exit Function
'------------------------------------------------------------------------------------------------------
SENSITIVITY_LINE:
'------------------------------------------------------------------------------------------------------
    TEMP_ARR = ASSET_OPEN_PRICE_UP_DOWN_FUNC(TICKER_STR, START_DATE, END_DATE, ii, PERCENT_VAL, , , , 2)
    If IsArray(TEMP_ARR) = False Then
        For j = jj To kk: TEMP_MATRIX(i, j) = "": Next j
        GoTo 1983
    End If
    k = LBound(TEMP_ARR)
    For j = jj To kk
        TEMP_MATRIX(i, j) = TEMP_ARR(k)
        k = k + 1
    Next j
1983:
'------------------------------------------------------------------------------------------------------
Return
'------------------------------------------------------------------------------------------------------
ERROR_LABEL:
ASSETS_OPEN_PRICE_UP_DOWN_FUNC = Err.number
End Function

'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
'BUY at the OPEN!!
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
'Yesterday morning I figured my favourite stock will open UP big time.
'that 's based upon exhaustive and very sophisticated mathematical
'analysis. Though a 3% increase in the Open (from the previous Close) would guarantee
'a gain (from Open to Close) of 0.68022% (accurate to five places of decimal), I go
'for the 2%. I'm not greedy, eh? GCE does open at 2% greater than the previous Close and
'I buy it at $4.12 ... eagerly. Then I watch in dismay as it drops to $4.06.
'Now I 'm investigating the error in the assumptions. Math is never wrong, right?
'Now there 's this Theory of Gremlins in Mathematics.
'After careful analysis, I predict my favourite stock will open UP big time, this morning.
'Have I said that before?
'-----------------------------
'A week ago I made a prediction (for my brother-in-law).
'I said that his CBQ = BRIC stock (then $27.75) would hit $31 in two weeks.
'Though we both laugh, he keeps reminding me of the prediction.
'I take a peek this morning and find that it's trading at (about) $30.
'If my predictions comes to pass, people may actually take 'em seriously
'... and that ain't good.
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------

Function ASSET_OPEN_PRICE_UP_DOWN_FUNC(ByVal TICKER_STR As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal END_DATE As Date, _
Optional ByVal VERSION As Integer = 0, _
Optional ByVal PERCENT_VAL As Variant = 3.52411575562702E-02, _
Optional ByVal MIN_PERCENT_VAL As Variant = 0#, _
Optional ByVal MAX_PERCENT_VAL As Variant = 0.05, _
Optional ByVal DELTA_PERCENT_VAL As Variant = 0.25 / 100, _
Optional ByVal OUTPUT As Integer = 3)

'VERSION Schemes:
'UP = 0  You BUY at the Open when it's greater than the previous Close by (at least) X%.
'DOWN = 1  You BUY at the Open when it's less than the previous Close by (at least) X%.

Dim i As Long
Dim j As Long
Dim k As Long

Dim ii As Double
Dim jj As Double
Dim kk As Double

Dim NROWS As Long
Dim NBINS As Long
Dim PERCENT_VAL As Variant

Dim TEMP1_SUM As Double
Dim TEMP2_SUM As Double
Dim TEMP3_SUM As Double

Dim CAGR_VAL As Double 'approximation
Dim MEAN_VAL As Double
Dim SIGMA_VAL As Double

Dim TEMP_STR As String

Dim SYMBOL_STR As String
Dim TEMP_ARR As Variant
Dim CONST_BOX As Variant
Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

'------------------------------------------------------------------------------------------------------
On Error GoTo ERROR_LABEL
'------------------------------------------------------------------------------------------------------
If IsArray(TICKER_STR) = False Then
    SYMBOL_STR = TICKER_STR
    DATA_MATRIX = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, START_DATE, END_DATE, _
                  "d", "DOHLCVA", False, True, True)
Else
    SYMBOL_STR = "ABC"
    DATA_MATRIX = TICKER_STR
End If
NROWS = UBound(DATA_MATRIX, 1)

If PERCENT_VAL = "" Then
    GoSub MAX_LINE
    PERCENT_VAL = PERCENT_VAL
End If
If OUTPUT > 3 Then
    ASSET_OPEN_PRICE_UP_DOWN_FUNC = PERCENT_VAL
    Exit Function
End If
'------------------------------------------------------------------------------------------------------
ReDim TEMP_MATRIX(0 To NROWS, 1 To 14)
'------------------------------------------------------------------------------------------------------
TEMP_MATRIX(0, 1) = "DATE"
TEMP_MATRIX(0, 2) = "OPEN"
TEMP_MATRIX(0, 3) = "HIGH"
TEMP_MATRIX(0, 4) = "LOW"
TEMP_MATRIX(0, 5) = "CLOSE"
TEMP_MATRIX(0, 6) = "VOLUME"
TEMP_MATRIX(0, 7) = "ADJ CLOSE"
TEMP_MATRIX(0, 8) = "RETURNS"
TEMP_MATRIX(0, 9) = "OPEN CHANGE"
TEMP_MATRIX(0, 10) = "CLOSE HIGHER"
TEMP_MATRIX(0, 11) = "INCREASE"
TEMP_MATRIX(0, 12) = "CLOSE LOWER"
TEMP_MATRIX(0, 13) = "DECREASE"
TEMP_MATRIX(0, 14) = "CHANGE"
'------------------------------------------------------------------------------------------------------
i = 1
For j = 1 To 7: TEMP_MATRIX(i, j) = DATA_MATRIX(i, j): Next j
TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 6) / 1000
TEMP_MATRIX(i, 8) = DATA_MATRIX(i, 5) / DATA_MATRIX(i, 2) - 1
For j = 9 To 14: TEMP_MATRIX(i, j) = "": Next j
'------------------------------------------------------------------------------------------------------
ii = 0: jj = 0: kk = 0
TEMP1_SUM = 0: TEMP2_SUM = 0: TEMP3_SUM = 0
'------------------------------------------------------------------------------------------------------
If VERSION = 0 Then
'------------------------------------------------------------------------------------------------------
    For i = 2 To NROWS
        For j = 1 To 7: TEMP_MATRIX(i, j) = DATA_MATRIX(i, j): Next j
        TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 6) / 1000
        TEMP_MATRIX(i, 8) = TEMP_MATRIX(i, 7) / TEMP_MATRIX(i - 1, 7) - 1
        If TEMP_MATRIX(i, 2) >= (1 + PERCENT_VAL) * TEMP_MATRIX(i - 1, 5) Then
            TEMP_MATRIX(i, 9) = 1
            ii = ii + 1
            TEMP_MATRIX(i, 14) = TEMP_MATRIX(i, 5) / TEMP_MATRIX(i, 2) - 1
            TEMP3_SUM = TEMP3_SUM + TEMP_MATRIX(i, 14)
        Else
            TEMP_MATRIX(i, 9) = ""
            TEMP_MATRIX(i, 14) = ""
        End If
        GoSub LOAD_LINE
    Next i
'------------------------------------------------------------------------------------------------------
Else
'------------------------------------------------------------------------------------------------------
    For i = 2 To NROWS
        For j = 1 To 7: TEMP_MATRIX(i, j) = DATA_MATRIX(i, j): Next j
        TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 6) / 1000
        TEMP_MATRIX(i, 8) = TEMP_MATRIX(i, 7) / TEMP_MATRIX(i - 1, 7) - 1
        If TEMP_MATRIX(i, 2) <= (1 + PERCENT_VAL) * TEMP_MATRIX(i - 1, 5) Then
            TEMP_MATRIX(i, 9) = 1
            ii = ii + 1
        
            TEMP_MATRIX(i, 14) = TEMP_MATRIX(i, 5) / TEMP_MATRIX(i, 2) - 1
            TEMP3_SUM = TEMP3_SUM + TEMP_MATRIX(i, 14)
        Else
            TEMP_MATRIX(i, 9) = ""
            TEMP_MATRIX(i, 14) = ""
        End If
        GoSub LOAD_LINE
    Next i
'------------------------------------------------------------------------------------------------------
End If
'------------------------------------------------------------------------------------------------------
If ii = 0 Then: ii = 10 ^ -5
If jj = 0 Then: jj = 10 ^ -5
If kk = 0 Then: kk = 10 ^ -5

'------------------------------------------------------------------------------------------------------
Select Case OUTPUT
'------------------------------------------------------------------------------------------------------
Case 0
'------------------------------------------------------------------------------------------------------
    ASSET_OPEN_PRICE_UP_DOWN_FUNC = TEMP_MATRIX
'------------------------------------------------------------------------------------------------------
Case 1
'------------------------------------------------------------------------------------------------------
    TEMP_STR = "If " & Format(SYMBOL_STR, "") & " opens " & IIf(VERSION = 0, "higher", "lower") & " by " & Format(PERCENT_VAL, "0.0%") & ", then Close > Open " & Format(jj / ii, "0.0%") & " of the time."
    TEMP_STR = TEMP_STR & " The Average (Close / Open) change is then " & Format(TEMP1_SUM / jj, "0.0%")
    TEMP_STR = TEMP_STR & ". It'll close " & IIf(VERSION = 0, "lower", "higher") & " " & Format(kk / ii, "0.0%") & " of the time: Avg (Close/Open) is " & Format(TEMP2_SUM / kk, "0.0%")

    ASSET_OPEN_PRICE_UP_DOWN_FUNC = TEMP_STR
'------------------------------------------------------------------------------------------------------
Case Else
'------------------------------------------------------------------------------------------------------
    MEAN_VAL = TEMP3_SUM / ii
    TEMP3_SUM = 0
    For i = 2 To NROWS
        If TEMP_MATRIX(i, 9) = 1 Then
            TEMP3_SUM = TEMP3_SUM + ((TEMP_MATRIX(i, 5) / TEMP_MATRIX(i, 2)) - 1 - MEAN_VAL) ^ 2
        End If
    Next i
    SIGMA_VAL = (TEMP3_SUM / ii) ^ 0.5
    CAGR_VAL = MEAN_VAL - 0.5 * (SIGMA_VAL) ^ 2 'approximation
    
    If OUTPUT = 2 Then
        ASSET_OPEN_PRICE_UP_DOWN_FUNC = Array(PERCENT_VAL, MEAN_VAL, SIGMA_VAL, CAGR_VAL, jj / ii, TEMP1_SUM / jj, kk / ii, TEMP2_SUM / kk)
    Else 'OUTPUT = 3
        NBINS = Int((MAX_PERCENT_VAL - MIN_PERCENT_VAL) / DELTA_PERCENT_VAL) + 1
        ReDim TEMP_MATRIX(0 To NBINS, 1 To 16)
        TEMP_MATRIX(0, 1) = " HIGHER: PERCENT_VAL"
        TEMP_MATRIX(0, 2) = " HIGHER: AVERAGE"
        TEMP_MATRIX(0, 3) = " HIGHER: VOLATILITY"
        TEMP_MATRIX(0, 4) = " HIGHER: CAGR"
        TEMP_MATRIX(0, 5) = " HIGHER: %C>O"
        TEMP_MATRIX(0, 6) = " HIGHER: AVERAGE(C>O)"
        TEMP_MATRIX(0, 7) = " HIGHER: %C<O"
        TEMP_MATRIX(0, 8) = " HIGHER: AVERAGE(C<O)"
    
        TEMP_MATRIX(0, 9) = " LOWER: PERCENT_VAL"
        TEMP_MATRIX(0, 10) = " LOWER: AVERAGE"
        TEMP_MATRIX(0, 11) = " LOWER: VOLATILITY"
        TEMP_MATRIX(0, 12) = " LOWER: CAGR"
        TEMP_MATRIX(0, 13) = " LOWER: %C>O"
        TEMP_MATRIX(0, 14) = " LOWER: AVERAGE(C>O)"
        TEMP_MATRIX(0, 15) = " LOWER: %C<O"
        TEMP_MATRIX(0, 16) = " LOWER: AVERAGE(C<O)"
        For j = 1 To 16: TEMP_MATRIX(0, j) = SYMBOL_STR & TEMP_MATRIX(0, j): Next j
        
        i = 1
        ii = 0: jj = 1: kk = 8
        PERCENT_VAL = ASSET_OPEN_PRICE_UP_DOWN_FUNC(DATA_MATRIX, , , ii, "", , , , 3) 'Max CAGR
        GoSub SENSITIVITY_LINE
        
        ii = 1: jj = 9: kk = 16
        PERCENT_VAL = ASSET_OPEN_PRICE_UP_DOWN_FUNC(DATA_MATRIX, , , ii, "", , , , 3) 'Max CAGR
        GoSub SENSITIVITY_LINE
        
        PERCENT_VAL = MIN_PERCENT_VAL
        For i = 2 To NBINS
            ii = 0: jj = 1: kk = 8
            GoSub SENSITIVITY_LINE
            ii = 1: jj = 9: kk = 16
            GoSub SENSITIVITY_LINE
            PERCENT_VAL = PERCENT_VAL + DELTA_PERCENT_VAL
        Next i
        ASSET_OPEN_PRICE_UP_DOWN_FUNC = TEMP_MATRIX
    End If
'------------------------------------------------------------------------------------------------------
End Select
'------------------------------------------------------------------------------------------------------
Exit Function
'------------------------------------------------------------------------------------------------------
LOAD_LINE:
'------------------------------------------------------------------------------------------------------
    If (TEMP_MATRIX(i, 9) = 1 And TEMP_MATRIX(i, 5) >= TEMP_MATRIX(i, 2)) Then
        TEMP_MATRIX(i, 10) = 1
        jj = jj + 1
    
        TEMP_MATRIX(i, 11) = TEMP_MATRIX(i, 5) / TEMP_MATRIX(i, 2) - 1
        TEMP1_SUM = TEMP1_SUM + TEMP_MATRIX(i, 11)
    Else
        TEMP_MATRIX(i, 10) = ""
        TEMP_MATRIX(i, 11) = ""
    End If
    
    If (TEMP_MATRIX(i, 9) = 1 And TEMP_MATRIX(i, 5) < TEMP_MATRIX(i, 2)) Then
        TEMP_MATRIX(i, 12) = 1
        kk = kk + 1
    
        TEMP_MATRIX(i, 13) = TEMP_MATRIX(i, 5) / TEMP_MATRIX(i, 2) - 1
        TEMP2_SUM = TEMP2_SUM + TEMP_MATRIX(i, 13)
    Else
        TEMP_MATRIX(i, 12) = ""
        TEMP_MATRIX(i, 13) = ""
    End If
'------------------------------------------------------------------------------------------------------
Return
'------------------------------------------------------------------------------------------------------
SENSITIVITY_LINE:
'------------------------------------------------------------------------------------------------------
    TEMP_ARR = ASSET_OPEN_PRICE_UP_DOWN_FUNC(DATA_MATRIX, , , ii, PERCENT_VAL, , , , 2)
    If IsArray(TEMP_ARR) = False Then
        For j = jj To kk: TEMP_MATRIX(i, j) = "": Next j
        GoTo 1983
    End If
    k = LBound(TEMP_ARR)
    For j = jj To kk
        TEMP_MATRIX(i, j) = TEMP_ARR(k)
        k = k + 1
    Next j
1983:
'------------------------------------------------------------------------------------------------------
Return
'------------------------------------------------------------------------------------------------------
MAX_LINE:
'------------------------------------------------------------------------------------------------------
    PUB_DATA_MATRIX = DATA_MATRIX
    PUB_VERSION = VERSION
    ReDim CONST_BOX(1 To 2, 1 To 1)
    CONST_BOX(1, 1) = MIN_PERCENT_VAL
    CONST_BOX(2, 1) = MAX_PERCENT_VAL
    PERCENT_VAL = UNIVAR_MIN_DIVIDE_CONQUER_FUNC("ASSET_OPEN_PRICE_UP_DOWN_OPTIMIZER_FUNC", CONST_BOX, False, , 600, 10 ^ -15)
'------------------------------------------------------------------------------------------------------
Return
'------------------------------------------------------------------------------------------------------
ERROR_LABEL:
'------------------------------------------------------------------------------------------------------
ASSET_OPEN_PRICE_UP_DOWN_FUNC = Err.number
End Function

Function ASSET_OPEN_PRICE_UP_DOWN_OPTIMIZER_FUNC(ByVal X_VAL As Double)

Dim i As Long
Dim j As Long
Dim NROWS As Long

Dim TEMP_SUM As Double
Dim MEAN_VAL As Double
Dim SIGMA_VAL As Double
Dim CAGR_VAL As Double

On Error GoTo ERROR_LABEL

NROWS = UBound(PUB_DATA_MATRIX, 1)
'------------------------------------------------------------------------------------------------------
TEMP_SUM = 0
'------------------------------------------------------------------------------------------------------
If PUB_VERSION = 0 Then
'------------------------------------------------------------------------------------------------------
    TEMP_SUM = 0
    j = 0
    For i = 2 To NROWS
        If PUB_DATA_MATRIX(i, 2) >= (1 + X_VAL) * PUB_DATA_MATRIX(i - 1, 5) Then
            TEMP_SUM = TEMP_SUM + PUB_DATA_MATRIX(i, 5) / PUB_DATA_MATRIX(i, 2) - 1
            j = j + 1
        End If
    Next i
    MEAN_VAL = TEMP_SUM / j
    TEMP_SUM = 0
    For i = 2 To NROWS
        If PUB_DATA_MATRIX(i, 2) >= (1 + X_VAL) * PUB_DATA_MATRIX(i - 1, 5) Then
            TEMP_SUM = TEMP_SUM + ((PUB_DATA_MATRIX(i, 5) / PUB_DATA_MATRIX(i, 2)) - 1 - MEAN_VAL) ^ 2
        End If
    Next i
'------------------------------------------------------------------------------------------------------
Else
'------------------------------------------------------------------------------------------------------
    TEMP_SUM = 0
    j = 0
    For i = 2 To NROWS
        If PUB_DATA_MATRIX(i, 2) <= (1 + X_VAL) * PUB_DATA_MATRIX(i - 1, 5) Then
            TEMP_SUM = TEMP_SUM + PUB_DATA_MATRIX(i, 5) / PUB_DATA_MATRIX(i, 2) - 1
            j = j + 1
        End If
    Next i
    MEAN_VAL = TEMP_SUM / j
    TEMP_SUM = 0
    For i = 2 To NROWS
        If PUB_DATA_MATRIX(i, 2) <= (1 + X_VAL) * PUB_DATA_MATRIX(i - 1, 5) Then
            TEMP_SUM = TEMP_SUM + ((PUB_DATA_MATRIX(i, 5) / PUB_DATA_MATRIX(i, 2)) - 1 - MEAN_VAL) ^ 2
        End If
    Next i
'------------------------------------------------------------------------------------------------------
End If
'------------------------------------------------------------------------------------------------------
SIGMA_VAL = (TEMP_SUM / j) ^ 0.5
CAGR_VAL = MEAN_VAL - 0.5 * (SIGMA_VAL) ^ 2 'approximation
'------------------------------------------------------------------------------------------------------

ASSET_OPEN_PRICE_UP_DOWN_OPTIMIZER_FUNC = CAGR_VAL 'Maximize

Exit Function
ERROR_LABEL:
ASSET_OPEN_PRICE_UP_DOWN_OPTIMIZER_FUNC = -2 ^ 52
End Function
