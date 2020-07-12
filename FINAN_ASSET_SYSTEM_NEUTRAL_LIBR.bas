Attribute VB_Name = "FINAN_ASSET_SYSTEM_NEUTRAL_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'Calculating Sharpe Ratio for Long-Only Versus Market-Neutral Strategies

Function ASSET_MARKET_NEUTRAL_SYSTEM_FUNC(ByVal TICKER_STR As Variant, _
ByVal INDEX_STR As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal END_DATE As Date, _
Optional ByVal CASH_RATE As Double = 0.04, _
Optional ByVal COUNT_BASIS As Double = 252, _
Optional ByVal OUTPUT As Integer = 1)

'INDEX_STR = SPY

Dim i As Long
Dim j As Long
Dim k As Long 'Max drawdown duration
Dim NROWS As Long

Dim LTEMP_SUM As Double
Dim NTEMP_SUM As Double

Dim LMEAN_VAL As Double
Dim NMEAN_VAL As Double

Dim LVOLAT_VAL As Double
Dim NVOLAT_VAL As Double

Dim LSHARPE_VAL As Double
Dim NSHARPE_VAL As Double
Dim MAX_DRAWDOWN As Double

Dim INDEX_VECTOR As Variant
Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

'----------------------------------------------------------------------------------------------------
If IsArray(TICKER_STR) = False Then
    DATA_MATRIX = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, START_DATE, END_DATE, _
                  "d", "DOHLCVA", False, True, True)
Else
    DATA_MATRIX = TICKER_STR
End If
'----------------------------------------------------------------------------------------------------
NROWS = UBound(DATA_MATRIX, 1)
'----------------------------------------------------------------------------------------------------
If IsArray(INDEX_STR) = False Then
    INDEX_VECTOR = YAHOO_HISTORICAL_DATA_SERIE_FUNC(INDEX_STR, START_DATE, END_DATE, _
                  "d", "A", False, False, True)
Else
    INDEX_VECTOR = INDEX_STR
End If
If UBound(INDEX_VECTOR, 1) <> NROWS Then: GoTo ERROR_LABEL
'----------------------------------------------------------------------------------------------------

ReDim TEMP_MATRIX(0 To NROWS, 1 To 16)
TEMP_MATRIX(0, 1) = "DATE"
TEMP_MATRIX(0, 2) = "OPEN"
TEMP_MATRIX(0, 3) = "HIGH"
TEMP_MATRIX(0, 4) = "LOW"
TEMP_MATRIX(0, 5) = "CLOSE"
TEMP_MATRIX(0, 6) = "VOLUME"
TEMP_MATRIX(0, 7) = "ADJ CLOSE"

TEMP_MATRIX(0, 8) = "DAILY RETURN"
TEMP_MATRIX(0, 9) = "EXCESS DAILY RETURN"

TEMP_MATRIX(0, 10) = "ADJ CLOSE INDEX"
TEMP_MATRIX(0, 11) = "DAILY RETURN INDEX"

TEMP_MATRIX(0, 12) = "NET DAILY RETURN"
TEMP_MATRIX(0, 13) = "CUMUL. RETURN"
TEMP_MATRIX(0, 14) = "HIGH WATERMARK"
TEMP_MATRIX(0, 15) = "DRAWDOWN"
TEMP_MATRIX(0, 16) = "MAX DRAWDOWN DURATION"

i = 1
For j = 1 To 7: TEMP_MATRIX(i, j) = DATA_MATRIX(i, j): Next j
TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 6) / 1000

TEMP_MATRIX(i, 8) = ""
TEMP_MATRIX(i, 9) = ""
TEMP_MATRIX(i, 10) = INDEX_VECTOR(i, 1)
For j = 11 To 16: TEMP_MATRIX(i, j) = "": Next j

i = 2
For j = 1 To 7: TEMP_MATRIX(i, j) = DATA_MATRIX(i, j): Next j
TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 6) / 1000
TEMP_MATRIX(i, 8) = TEMP_MATRIX(i, 7) / TEMP_MATRIX(i - 1, 7) - 1
TEMP_MATRIX(i, 9) = TEMP_MATRIX(i, 8) - CASH_RATE / COUNT_BASIS

TEMP_MATRIX(i, 10) = INDEX_VECTOR(i, 1)
TEMP_MATRIX(i, 11) = TEMP_MATRIX(i, 10) / TEMP_MATRIX(i - 1, 10) - 1

TEMP_MATRIX(i, 12) = (TEMP_MATRIX(i, 8) - TEMP_MATRIX(i, 11)) * 0.5
TEMP_MATRIX(i, 13) = TEMP_MATRIX(i, 12)
TEMP_MATRIX(i, 14) = TEMP_MATRIX(i, 13)
TEMP_MATRIX(i, 15) = (1 + TEMP_MATRIX(i, 14)) / (1 + TEMP_MATRIX(i, 13)) - 1
TEMP_MATRIX(i, 16) = IIf(TEMP_MATRIX(i, 15) = 0, 0, 0 + 1)

MAX_DRAWDOWN = TEMP_MATRIX(i, 15)
k = TEMP_MATRIX(i, 16)
LTEMP_SUM = TEMP_MATRIX(i, 9)
NTEMP_SUM = TEMP_MATRIX(i, 12)

For i = 3 To NROWS
    For j = 1 To 7: TEMP_MATRIX(i, j) = DATA_MATRIX(i, j): Next j
    TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 6) / 1000
    TEMP_MATRIX(i, 8) = TEMP_MATRIX(i, 7) / TEMP_MATRIX(i - 1, 7) - 1
    TEMP_MATRIX(i, 9) = TEMP_MATRIX(i, 8) - CASH_RATE / COUNT_BASIS
    
    TEMP_MATRIX(i, 10) = INDEX_VECTOR(i, 1)
    TEMP_MATRIX(i, 11) = TEMP_MATRIX(i, 10) / TEMP_MATRIX(i - 1, 10) - 1
    
    TEMP_MATRIX(i, 12) = (TEMP_MATRIX(i, 8) - TEMP_MATRIX(i, 11)) * 0.5
    TEMP_MATRIX(i, 13) = (1 + TEMP_MATRIX(i - 1, 13)) * (1 + TEMP_MATRIX(i, 12)) - 1
    
    If TEMP_MATRIX(i - 1, 14) > TEMP_MATRIX(i, 13) Then
        TEMP_MATRIX(i, 14) = TEMP_MATRIX(i - 1, 14)
    Else
        TEMP_MATRIX(i, 14) = TEMP_MATRIX(i, 13)
    End If
    TEMP_MATRIX(i, 15) = (1 + TEMP_MATRIX(i, 14)) / (1 + TEMP_MATRIX(i, 13)) - 1
    TEMP_MATRIX(i, 16) = IIf(TEMP_MATRIX(i, 15) = 0, 0, TEMP_MATRIX(i - 1, 16) + 1)
    
    If TEMP_MATRIX(i, 15) > MAX_DRAWDOWN Then: MAX_DRAWDOWN = TEMP_MATRIX(i, 15)
    If TEMP_MATRIX(i, 16) > k Then: k = TEMP_MATRIX(i, 16)
    
    LTEMP_SUM = LTEMP_SUM + TEMP_MATRIX(i, 9)
    NTEMP_SUM = NTEMP_SUM + TEMP_MATRIX(i, 12)
Next i

'---------------------------------------------------------------------------------------------------
Select Case OUTPUT
'---------------------------------------------------------------------------------------------------
Case 0
'---------------------------------------------------------------------------------------------------
    ASSET_MARKET_NEUTRAL_SYSTEM_FUNC = TEMP_MATRIX
'---------------------------------------------------------------------------------------------------
Case Else
'---------------------------------------------------------------------------------------------------
    LMEAN_VAL = LTEMP_SUM / (NROWS - 1)
    NMEAN_VAL = NTEMP_SUM / (NROWS - 1)
    
    LTEMP_SUM = 0: NTEMP_SUM = 0
    For i = 2 To NROWS
        LTEMP_SUM = LTEMP_SUM + (TEMP_MATRIX(i, 9) - LMEAN_VAL) ^ 2
        NTEMP_SUM = NTEMP_SUM + (TEMP_MATRIX(i, 12) - NMEAN_VAL) ^ 2
    Next i
    LVOLAT_VAL = (LTEMP_SUM / (NROWS - 2)) ^ 0.5 'sample
    NVOLAT_VAL = (NTEMP_SUM / (NROWS - 2)) ^ 0.5 'sample
        
    LSHARPE_VAL = LMEAN_VAL / LVOLAT_VAL * COUNT_BASIS ^ 0.5
    NSHARPE_VAL = NMEAN_VAL / NVOLAT_VAL * COUNT_BASIS ^ 0.5
    ASSET_MARKET_NEUTRAL_SYSTEM_FUNC = Array(LSHARPE_VAL, NSHARPE_VAL, MAX_DRAWDOWN, k)
'---------------------------------------------------------------------------------------------------
End Select
'---------------------------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
ASSET_MARKET_NEUTRAL_SYSTEM_FUNC = Err.number
End Function
