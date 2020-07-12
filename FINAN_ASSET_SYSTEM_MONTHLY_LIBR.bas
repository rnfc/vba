Attribute VB_Name = "FINAN_ASSET_SYSTEM_MONTHLY_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

Function ASSET_MONTHLY_SYSTEM_FUNC(ByVal TICKER_STR As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal END_DATE As Date, _
Optional ByVal BUY_MONTH As Long = 10, _
Optional ByVal SELL_MONTH As Long = 5, _
Optional ByVal INITIAL_CASH As Double = 1000000, _
Optional ByVal OUTPUT As Integer = 0)

Dim i As Long
Dim j As Long
Dim k As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

'Dim NO_DAYS As Long

Dim MEAN_VAL As Double
Dim VOLAT_VAL As Double
Dim SCAGR_VAL As Double
Dim XCAGR_VAL As Double

Dim HEADINGS_STR As String
Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

If IsArray(TICKER_STR) = False Then
    DATA_MATRIX = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, START_DATE, END_DATE, "MONTHLY", "DOHLCVA", False, True, True)
Else
    DATA_MATRIX = TICKER_STR
End If
NROWS = UBound(DATA_MATRIX, 1)
'-----------------------------------------------------------------------------------------------------
HEADINGS_STR = "DATE,OPEN,HIGH,LOW,CLOSE,VOLUME,ADJ.CLOSE,RETURN,MONTH,EQUITY,CASH,SYSTEM BALANCE,BUY HOLD BALANCE,"
j = Len(HEADINGS_STR)
NCOLUMNS = 0
For i = 1 To j
    If Mid(HEADINGS_STR, i, 1) = "," Then: NCOLUMNS = NCOLUMNS + 1
Next i
ReDim TEMP_MATRIX(0 To NROWS, 1 To NCOLUMNS)
i = 1
For k = 1 To NCOLUMNS
    j = InStr(i, HEADINGS_STR, ",")
    TEMP_MATRIX(0, k) = Mid(HEADINGS_STR, i, j - i)
    i = j + 1
Next k
'-----------------------------------------------------------------------------------------------------

i = 1
For j = 1 To 7: TEMP_MATRIX(i, j) = DATA_MATRIX(i, j): Next j
TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 6) / 1000
TEMP_MATRIX(i, 8) = TEMP_MATRIX(i, 5) / TEMP_MATRIX(i, 2) - 1
TEMP_MATRIX(i, 9) = Month(TEMP_MATRIX(i, 1))
TEMP_MATRIX(i, 10) = 0
TEMP_MATRIX(i, 11) = INITIAL_CASH
TEMP_MATRIX(i, 12) = TEMP_MATRIX(i, 10) + TEMP_MATRIX(i, 11)
TEMP_MATRIX(i, 13) = TEMP_MATRIX(i, 10) + TEMP_MATRIX(i, 11)

For i = 2 To NROWS
    For j = 1 To 7: TEMP_MATRIX(i, j) = DATA_MATRIX(i, j): Next j
    TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 6) / 1000
    TEMP_MATRIX(i, 8) = TEMP_MATRIX(i, 7) / TEMP_MATRIX(i - 1, 7) - 1
    k = ((1 + TEMP_MATRIX(i - 1, 9)) Mod 13)
    TEMP_MATRIX(i, 9) = IIf(k > 1, k, 1) '(TEMP_MATRIX(i - 1, 9) Mod 12) + 1 'Month(TEMP_MATRIX(i, 1))
    
    If TEMP_MATRIX(i, 9) = SELL_MONTH Then
        TEMP_MATRIX(i, 10) = 0
    Else
        If TEMP_MATRIX(i, 9) = BUY_MONTH Then
            TEMP_MATRIX(i, 10) = TEMP_MATRIX(i - 1, 11)
        Else
            TEMP_MATRIX(i, 10) = TEMP_MATRIX(i - 1, 10) * (1 + TEMP_MATRIX(i, 8))
        End If
    End If
    
    If (TEMP_MATRIX(i - 1, 10) > 0 And TEMP_MATRIX(i, 10) = 0) Then
        TEMP_MATRIX(i, 11) = TEMP_MATRIX(i - 1, 10) * (1 + TEMP_MATRIX(i, 8))
    Else
        If (TEMP_MATRIX(i - 1, 10) = 0 And TEMP_MATRIX(i, 10) > 0) Then
            TEMP_MATRIX(i, 11) = 0
        Else
            TEMP_MATRIX(i, 11) = TEMP_MATRIX(i - 1, 11)
        End If
    End If
    
    TEMP_MATRIX(i, 12) = TEMP_MATRIX(i, 10) + TEMP_MATRIX(i, 11)
    TEMP_MATRIX(i, 13) = (TEMP_MATRIX(i, 8) + 1) * TEMP_MATRIX(i - 1, 13)
    MEAN_VAL = MEAN_VAL + (TEMP_MATRIX(i, 12) / TEMP_MATRIX(i - 1, 12) - 1)
Next i

'NO_DAYS = TEMP_MATRIX(NROWS, 1) - TEMP_MATRIX(1, 1)
SCAGR_VAL = (TEMP_MATRIX(NROWS, 12) / TEMP_MATRIX(1, 12)) ^ (12 / NROWS) - 1 '(365 / NO_DAYS) - 1
XCAGR_VAL = (TEMP_MATRIX(NROWS, 13) / TEMP_MATRIX(1, 13)) ^ (12 / NROWS) - 1 '(365 / NO_DAYS) - 1

TEMP_MATRIX(0, 12) = TEMP_MATRIX(0, 12) & ": CAGR = " & Format(SCAGR_VAL, "0.00%")
TEMP_MATRIX(0, 13) = TEMP_MATRIX(0, 13) & ": CAGR = " & Format(XCAGR_VAL, "0.00%")

If OUTPUT = 0 Then
    ASSET_MONTHLY_SYSTEM_FUNC = TEMP_MATRIX
Else
    MEAN_VAL = MEAN_VAL / (NROWS - 1)
    
    For i = 2 To NROWS
        VOLAT_VAL = VOLAT_VAL + ((TEMP_MATRIX(i, 12) / TEMP_MATRIX(i - 1, 12) - 1) - MEAN_VAL) ^ 2
    Next i
    VOLAT_VAL = (VOLAT_VAL / (NROWS - 1)) ^ 0.5
    
    If OUTPUT = 1 Then
        ASSET_MONTHLY_SYSTEM_FUNC = MEAN_VAL / VOLAT_VAL
    Else
        If OUTPUT = 2 Then
            ASSET_MONTHLY_SYSTEM_FUNC = Array(MEAN_VAL / VOLAT_VAL, MEAN_VAL, VOLAT_VAL, SCAGR_VAL, XCAGR_VAL)
        Else
            ASSET_MONTHLY_SYSTEM_FUNC = Array(SCAGR_VAL, XCAGR_VAL)
        End If
    End If
End If

Exit Function
ERROR_LABEL:
ASSET_MONTHLY_SYSTEM_FUNC = "--"
End Function

Function ASSETS_MONTHLY_SYSTEM_FUNC(ByRef TICKERS_RNG As Variant, _
ByVal START_DATE As Date, _
ByVal END_DATE As Date, _
Optional ByVal BUY_MONTH As Long = 0, _
Optional ByVal SELL_MONTH As Long = 0, _
Optional ByVal INITIAL_CASH As Double = 5000, _
Optional ByVal VERSION As Integer = 0)

Dim g As Long
Dim h As Long

Dim i As Long
Dim j As Long

Dim ii As Long
Dim jj As Long

Dim iii As Long
Dim jjj As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

'Dim NO_DAYS As Long

Dim SCAGR_VAL As Double
Dim XCAGR_VAL As Double

Dim DATE0_VAL As Date
Dim DATE1_VAL As Date
Dim DATE2_VAL As Date

Dim PRICE0_VAL As Double
Dim PRICE1_VAL As Double
Dim PRICE2_VAL As Double

Dim TICKER_STR As String

Dim CASH1_VAL As Double
Dim CASH2_VAL As Double

Dim EQUITY1_VAL As Double
Dim EQUITY2_VAL As Double

Dim BUY_HOLD0_VAL As Double
Dim BUY_HOLD1_VAL As Double
Dim BUY_HOLD2_VAL As Double

Dim SYSTEM0_VAL As Double
Dim SYSTEM1_VAL As Double
Dim SYSTEM2_VAL As Double

Dim TEMP_VAL As Double
Dim TEMP_RETURN As Double

Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant
Dim TICKERS_VECTOR As Variant

Dim MEAN_VAL As Double
Dim SIGMA_VAL As Double
Dim DATA_ARR() As Double

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
NCOLUMNS = UBound(TICKERS_VECTOR, 1)

ReDim TEMP_MATRIX(0 To NCOLUMNS, 1 To 10)
TEMP_MATRIX(0, 1) = "TICKER"
TEMP_MATRIX(0, 2) = "STARTING PERIOD"
TEMP_MATRIX(0, 3) = "ENDING PERIOD"
TEMP_MATRIX(0, 4) = "BUY MONTH"
TEMP_MATRIX(0, 5) = "SELL MONTH"
TEMP_MATRIX(0, 6) = "SYSTEM CAGR"
TEMP_MATRIX(0, 7) = "BUY HOLD CAGR"
TEMP_MATRIX(0, 8) = "SYSTEM MONTHLY AVG RETURN"
TEMP_MATRIX(0, 9) = "SYSTEM MONTHLY VOLATILITY"
TEMP_MATRIX(0, 10) = "SYSTEM SHARPE"

'-------------------------------------------------------------------------------------------------------
For j = 1 To NCOLUMNS
'-------------------------------------------------------------------------------------------------------

    TICKER_STR = TICKERS_VECTOR(j, 1)
    TEMP_MATRIX(j, 1) = TICKER_STR
    
    DATA_MATRIX = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, START_DATE, END_DATE, "m", "DA", False, False, True)
    If IsArray(DATA_MATRIX) = False Then: GoTo 1983
    NROWS = UBound(DATA_MATRIX, 1)
    If NROWS = 1 Then: GoTo 1983
    
    If BUY_MONTH = 0 Or SELL_MONTH = 0 Then
        iii = 1: jjj = 12
        BUY_MONTH = iii: SELL_MONTH = jjj
        GoSub MODEL_LINE
        If VERSION = 0 Then 'Max Growth
            TEMP_VAL = SCAGR_VAL
            For ii = 1 To 12 'buy month
                For jj = 1 To 12 'sell month
                    BUY_MONTH = ii
                    SELL_MONTH = jj
                    GoSub MODEL_LINE
                    If SCAGR_VAL > TEMP_VAL Then
                        TEMP_VAL = SCAGR_VAL
                        iii = ii
                        jjj = jj
                    End If
                Next jj
            Next ii
        Else 'Max Sharpe
            TEMP_VAL = MEAN_VAL / SIGMA_VAL
            For ii = 1 To 12 'buy month
                For jj = 1 To 12 'sell month
                    BUY_MONTH = ii
                    SELL_MONTH = jj
                    GoSub MODEL_LINE
                    If (MEAN_VAL / SIGMA_VAL) > TEMP_VAL Then
                        TEMP_VAL = MEAN_VAL / SIGMA_VAL
                        iii = ii
                        jjj = jj
                    End If
                Next jj
            Next ii
        End If
        BUY_MONTH = iii
        SELL_MONTH = jjj
        GoSub MODEL_LINE
        BUY_MONTH = 0
        SELL_MONTH = 0
    Else
        GoSub MODEL_LINE
        iii = BUY_MONTH
        jjj = SELL_MONTH
    End If

    TEMP_MATRIX(j, 2) = DATE0_VAL
    TEMP_MATRIX(j, 3) = DATE1_VAL
    TEMP_MATRIX(j, 4) = Format(DateSerial(Year(Now), iii, 1), "mmm dd")
    TEMP_MATRIX(j, 5) = Format(DateSerial(Year(Now), jjj, 1), "mmm dd")
    TEMP_MATRIX(j, 6) = SCAGR_VAL
    TEMP_MATRIX(j, 7) = XCAGR_VAL
    TEMP_MATRIX(j, 8) = MEAN_VAL
    TEMP_MATRIX(j, 9) = SIGMA_VAL
    TEMP_MATRIX(j, 10) = MEAN_VAL / SIGMA_VAL
    
1983:
'-------------------------------------------------------------------------------------------------------
Next j
'-------------------------------------------------------------------------------------------------------

ASSETS_MONTHLY_SYSTEM_FUNC = TEMP_MATRIX

Exit Function
'-----------------------------------------------------------------------------------------------------------
MODEL_LINE:
'-----------------------------------------------------------------------------------------------------------
    
    i = 1
    DATE0_VAL = DATA_MATRIX(i, 1)
    DATE1_VAL = DATE0_VAL
    g = Month(DATE1_VAL)
    
    PRICE0_VAL = DATA_MATRIX(i, 2)
    PRICE1_VAL = PRICE0_VAL
    
    CASH1_VAL = INITIAL_CASH
    EQUITY1_VAL = 0
    
    BUY_HOLD0_VAL = INITIAL_CASH
    BUY_HOLD1_VAL = BUY_HOLD0_VAL
    
    SYSTEM0_VAL = INITIAL_CASH
    SYSTEM1_VAL = SYSTEM0_VAL
    
    MEAN_VAL = 0
    ReDim DATA_ARR(2 To NROWS)
    For i = 2 To NROWS
'-----------------------------------------------------------------------------------------------------------
        h = ((1 + g) Mod 13)
        h = IIf(h > 1, h, 1)
        DATE2_VAL = DATA_MATRIX(i, 1)
        PRICE2_VAL = DATA_MATRIX(i, 2)
        TEMP_RETURN = PRICE2_VAL / PRICE1_VAL - 1
        
        If h = SELL_MONTH Then
            EQUITY2_VAL = 0
        Else
            If h = BUY_MONTH Then
                EQUITY2_VAL = CASH1_VAL
            Else
                EQUITY2_VAL = EQUITY1_VAL * (1 + TEMP_RETURN)
            End If
        End If
        
        If (EQUITY1_VAL > 0 And EQUITY2_VAL = 0) Then
            CASH2_VAL = EQUITY1_VAL * (1 + TEMP_RETURN)
        Else
            If (EQUITY1_VAL = 0 And EQUITY2_VAL > 0) Then
                CASH2_VAL = 0
            Else
                CASH2_VAL = CASH1_VAL
            End If
        End If
    
        SYSTEM2_VAL = CASH2_VAL + EQUITY2_VAL
        DATA_ARR(i) = SYSTEM2_VAL / SYSTEM1_VAL - 1
        MEAN_VAL = MEAN_VAL + DATA_ARR(i)
        
        BUY_HOLD2_VAL = BUY_HOLD1_VAL * (1 + TEMP_RETURN)
        
        DATE1_VAL = DATE2_VAL
        g = Month(DATE1_VAL)
        
        PRICE1_VAL = PRICE2_VAL
        
        CASH1_VAL = CASH2_VAL
        EQUITY1_VAL = EQUITY2_VAL
        
        SYSTEM1_VAL = SYSTEM2_VAL
        BUY_HOLD1_VAL = BUY_HOLD2_VAL
'-----------------------------------------------------------------------------------------------------------
    Next i
'-----------------------------------------------------------------------------------------------------------
    MEAN_VAL = MEAN_VAL / (NROWS - 1)
    SIGMA_VAL = 0
    For i = 2 To NROWS
        SIGMA_VAL = SIGMA_VAL + (DATA_ARR(i) - MEAN_VAL) ^ 2
    Next i
    SIGMA_VAL = (SIGMA_VAL / (NROWS - 1)) ^ 0.5
    If SIGMA_VAL = 0 Then: SIGMA_VAL = 10 ^ -5
    'NO_DAYS = DATE1_VAL - DATE0_VAL
    'If NO_DAYS = 0 Then: NO_DAYS = 1
    
    SCAGR_VAL = (SYSTEM1_VAL / SYSTEM0_VAL) ^ (12 / NROWS) - 1 '(365 / NO_DAYS) - 1
    XCAGR_VAL = (BUY_HOLD1_VAL / BUY_HOLD0_VAL) ^ (12 / NROWS) - 1 '(365 / NO_DAYS) - 1

'-----------------------------------------------------------------------------------------------------------
Return
'-----------------------------------------------------------------------------------------------------------
ERROR_LABEL:
ASSETS_MONTHLY_SYSTEM_FUNC = Err.number
End Function

