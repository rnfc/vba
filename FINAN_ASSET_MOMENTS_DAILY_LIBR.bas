Attribute VB_Name = "FINAN_ASSET_MOMENTS_DAILY_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : ASSETS_BEST_WORST_ROC_FUNC
'DESCRIPTION   : Worst, Best and Average Return on Capital
'LIBRARY       : FINAN_ASSET
'GROUP         : MOMENTS
'ID            : 011

'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function ASSETS_BEST_WORST_ROC_FUNC(ByRef TICKERS_RNG As Variant, _
ByVal START_DATE As Date, _
ByVal END_DATE As Date, _
Optional ByRef PERIODS As Long = 10)

Dim i As Long
Dim k As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_VAL As Double
Dim MIN_VAL As Double
Dim MAX_VAL As Double
Dim TEMP_SUM As Double
Dim TICKER_STR As String
Dim DATA_VECTOR As Variant
Dim TICKERS_VECTOR As Variant

On Error GoTo ERROR_LABEL

If PERIODS = 0 Then: PERIODS = 10

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

ReDim TEMP_MATRIX(0 To NCOLUMNS, 1 To 4)
TEMP_MATRIX(0, 1) = "SYMBOL"
TEMP_MATRIX(0, 2) = "WORST ROC"
TEMP_MATRIX(0, 3) = "BEST ROC"
TEMP_MATRIX(0, 4) = "AVERAGE ROC"

For k = 1 To NCOLUMNS
    TICKER_STR = TICKERS_VECTOR(k, 1)
    TEMP_MATRIX(k, 1) = TICKER_STR
    DATA_VECTOR = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, START_DATE, END_DATE, "DAILY", "A", False, False, False)
    If IsArray(DATA_VECTOR) = False Then: GoTo 1983
    NROWS = UBound(DATA_VECTOR, 1)
    If (NROWS - PERIODS) <= 0 Then: GoTo 1983
    MIN_VAL = 2 ^ 52: MAX_VAL = -2 ^ 52
    TEMP_SUM = 0
    For i = 1 To NROWS - PERIODS
        TEMP_VAL = DATA_VECTOR(i + PERIODS, 1) / DATA_VECTOR(i, 1) - 1
        TEMP_SUM = TEMP_SUM + TEMP_VAL
        If TEMP_VAL > MAX_VAL Then: MAX_VAL = TEMP_VAL
        If TEMP_VAL < MIN_VAL Then: MIN_VAL = TEMP_VAL
    Next i
    TEMP_MATRIX(k, 2) = MIN_VAL 'Worst
    TEMP_MATRIX(k, 3) = MAX_VAL 'Best
    TEMP_MATRIX(k, 4) = TEMP_SUM / (NROWS - PERIODS) 'Average
1983:
Next k

ASSETS_BEST_WORST_ROC_FUNC = TEMP_MATRIX


Exit Function
ERROR_LABEL:
ASSETS_BEST_WORST_ROC_FUNC = Err.number
End Function


'Percentage of days having 1-day changes in the ranges indicated

Function ASSET_DAILY_RETURNS_HISTOGRAM_FUNC(ByRef TICKER_STR As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal END_DATE As Date, _
Optional ByVal MIN_BIN As Double = -0.05, _
Optional ByVal DELTA_BIN As Double = 0.01, _
Optional ByVal NBINS As Long = 12)

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim TEMP_BIN As Double
Dim TEMP_RETURN As Double

Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

If IsArray(TICKER_STR) = True Then
    DATA_MATRIX = TICKER_STR
Else
    DATA_MATRIX = _
    YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, START_DATE, END_DATE, "d", "DOHLCVA", False, True, True)
End If
NROWS = UBound(DATA_MATRIX, 1)

ReDim TEMP_MATRIX(0 To NROWS + 1, 1 To NBINS + 2)

TEMP_MATRIX(0, 1) = "Date": TEMP_MATRIX(0, 2) = "Return"
TEMP_MATRIX(1, 1) = "": TEMP_MATRIX(1, 2) = ""
For i = 1 To NROWS
    If i <> 1 Then
        TEMP_RETURN = DATA_MATRIX(i, 5) / DATA_MATRIX(i - 1, 5) - 1
    Else
        TEMP_RETURN = DATA_MATRIX(i, 5) / DATA_MATRIX(i, 2) - 1
    End If
    TEMP_MATRIX(i + 1, 1) = DATA_MATRIX(i, 1)
    TEMP_MATRIX(i + 1, 2) = TEMP_RETURN

    TEMP_BIN = MIN_BIN
    For j = 1 To NBINS
        If j = 1 Then
            TEMP_MATRIX(0, j + 2) = "<" & Format(TEMP_BIN, "0.00%")
            TEMP_MATRIX(i + 1, j + 2) = IIf(TEMP_RETURN < TEMP_BIN, 1, 0)
        ElseIf j = NBINS Then
            TEMP_MATRIX(0, j + 2) = ">" & Format((TEMP_BIN - DELTA_BIN), "0.00%")
            TEMP_MATRIX(i + 1, j + 2) = IIf(TEMP_RETURN > (TEMP_BIN - DELTA_BIN), 1, 0)
        Else
            TEMP_MATRIX(0, j + 2) = Format((TEMP_BIN - DELTA_BIN), "0.00%") & ", " & Format(TEMP_BIN, "0.00%")
            TEMP_MATRIX(i + 1, j + 2) = IIf((TEMP_RETURN > (TEMP_BIN - DELTA_BIN) And TEMP_RETURN <= TEMP_BIN), 1, 0)
        End If
        TEMP_MATRIX(1, j + 2) = TEMP_MATRIX(1, j + 2) + TEMP_MATRIX(i + 1, j + 2)
        TEMP_BIN = TEMP_BIN + DELTA_BIN
    Next j
Next i

For j = 1 To NBINS
    TEMP_MATRIX(1, j + 2) = TEMP_MATRIX(1, j + 2) / NROWS
Next j
ASSET_DAILY_RETURNS_HISTOGRAM_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
ASSET_DAILY_RETURNS_HISTOGRAM_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : ASSET_COMPOUND_DAILY_RETURNS_FUNC
'DESCRIPTION   : Weekly & Monthlys Returns
'LIBRARY       : FINAN_ASSET
'GROUP         : MOMENTS
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function ASSET_COMPOUND_DAILY_RETURNS_FUNC(ByVal TICKER_STR As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal END_DATE As Date, _
Optional ByVal OUTPUT As Integer = 0)

Dim i As Long
Dim j As Long
Dim k As Long
Dim h(1 To 12) As Long
Dim m(1 To 12) As Double
Dim s(1 To 12) As Double

Dim NROWS As Long
Dim RETURNS_VECTOR As Variant

Dim WEEKS_MATRIX As Variant
Dim UPS_MATRIX As Variant
Dim MONTHS_MATRIX As Variant
Dim DAILY_MATRIX As Variant
Dim SUMMARY_MATRIX As Variant

Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

If IsArray(TICKER_STR) = True Then
    DATA_MATRIX = TICKER_STR
Else
    DATA_MATRIX = _
    YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, START_DATE, END_DATE, "D", "DA", True, False, True)
End If
NROWS = UBound(DATA_MATRIX, 1)

ReDim RETURNS_VECTOR(1 To NROWS - 1, 1 To 1)
ReDim WEEKS_MATRIX(0 To NROWS, 1 To 5)
ReDim UPS_MATRIX(0 To NROWS, 1 To 5)
ReDim MONTHS_MATRIX(0 To NROWS, 1 To 12)
ReDim DAILY_MATRIX(0 To NROWS, 1 To 12)

WEEKS_MATRIX(0, 1) = "MON"
WEEKS_MATRIX(0, 2) = "TUE"
WEEKS_MATRIX(0, 3) = "WED"
WEEKS_MATRIX(0, 4) = "THU"
WEEKS_MATRIX(0, 5) = "FRID"
For i = 1 To 5: UPS_MATRIX(0, i) = WEEKS_MATRIX(0, i): Next i

For i = 1 To 12: s(i) = 0: Next i
For i = 2 To NROWS
    RETURNS_VECTOR(i - 1, 1) = DATA_MATRIX(i, 2) / DATA_MATRIX(i - 1, 2) - 1
'----------------------FIRST PASS: Identify the Day of the Week-----------------
    For j = 1 To 5
        WEEKS_MATRIX(i, j) = IIf(Weekday(DATA_MATRIX(i, 1)) = j + 1, 1, "")
        If WEEKS_MATRIX(i, j) <> "" Then: s(j) = s(j) + WEEKS_MATRIX(i, j)
    Next j

    If DATA_MATRIX(i, 2) > DATA_MATRIX(i - 1, 2) Then
        k = 1
    Else
        k = 0
    End If
'----------------------SECOND PASS: Identify UP days----------------------------
'--------------------------UP = "1", DN = "0"-----------------------------------
    For j = 1 To 5
        UPS_MATRIX(i, j) = IIf(WEEKS_MATRIX(i, j) = 1, k, "")
        If UPS_MATRIX(i, j) <> "" Then: s(5 + j) = s(5 + j) + UPS_MATRIX(i, j)
    Next j
'-------------------------------------------------------------------------------
Next i

For j = 1 To 5
    If s(j) <> 0 Then: UPS_MATRIX(1, j) = s(5 + j) / s(j)
Next j

'-------------THIRD PASS: Identify Daily Returns for each Month-----------------
'-------------FORTH PASS: Identify Compound Daily Returns-----------------------

For i = 1 To 12
    MONTHS_MATRIX(0, i) = UCase(Format(DateSerial(Year(Date), i, Day(Date)), "mmm"))
    DAILY_MATRIX(0, i) = MONTHS_MATRIX(0, i)
Next i

For i = 1 To 12
    h(i) = 0
    m(i) = 1
    s(i) = 0
Next i
 
For i = 2 To NROWS
    For j = 1 To 12
        MONTHS_MATRIX(i, j) = IIf(Month(DATA_MATRIX(i, 1)) = j, RETURNS_VECTOR(i - 1, 1), "")
        If MONTHS_MATRIX(i, j) <> "" Then
            s(j) = s(j) + MONTHS_MATRIX(i, j)
            h(j) = h(j) + 1
            DAILY_MATRIX(i, j) = IIf(Month(DATA_MATRIX(i, 1)) = j, 1 + MONTHS_MATRIX(i, j), 1)
        Else
            DAILY_MATRIX(i, j) = 1
        End If
        m(j) = m(j) * DAILY_MATRIX(i, j)
    Next j
Next i

For j = 1 To 12
    If h(j) <> 0 Then
        MONTHS_MATRIX(1, j) = s(j) / h(j)
        DAILY_MATRIX(1, j) = m(j) ^ (1 / h(j)) - 1
    End If
Next j
'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------
ReDim SUMMARY_MATRIX(1 To 5, 1 To 13)

For i = 1 To 5: SUMMARY_MATRIX(1, i + 1) = WEEKS_MATRIX(0, i): Next i
For i = 7 To 13: SUMMARY_MATRIX(1, i) = "": Next i
For i = 1 To 12: SUMMARY_MATRIX(3, i + 1) = MONTHS_MATRIX(0, i): Next i
SUMMARY_MATRIX(1, 1) = "DAY:"
SUMMARY_MATRIX(2, 1) = "PERCENTAGE UP DAYS"
SUMMARY_MATRIX(3, 1) = "MONTH:"
SUMMARY_MATRIX(4, 1) = "AVERAGE DAILY GROWTH RATE"
SUMMARY_MATRIX(5, 1) = "COMPOUND DAILY GROWTH RATE"

For i = 1 To 12
    If i < 6 Then
        SUMMARY_MATRIX(2, i + 1) = UPS_MATRIX(1, i)
    Else
        SUMMARY_MATRIX(2, i + 1) = ""
    End If
            
    SUMMARY_MATRIX(4, i + 1) = MONTHS_MATRIX(1, i)
    SUMMARY_MATRIX(5, i + 1) = DAILY_MATRIX(1, i)
Next i
        
'-----------------------------------------------------------------------------
Select Case OUTPUT
'-----------------------------------------------------------------------------
Case 0
'-----------------------------------------------------------------------------
    ASSET_COMPOUND_DAILY_RETURNS_FUNC = SUMMARY_MATRIX 'Summary
'-----------------------------------------------------------------------------
Case 1
'-----------------------------------------------------------------------------
    ASSET_COMPOUND_DAILY_RETURNS_FUNC = DAILY_MATRIX 'Identify Compound
    'Daily Returns assuming you get ONLY returns for Jan or Feb etc.
'-----------------------------------------------------------------------------
Case 2
'-----------------------------------------------------------------------------
    ASSET_COMPOUND_DAILY_RETURNS_FUNC = MONTHS_MATRIX
    'Identify Daily Returns for each Month
'-----------------------------------------------------------------------------
Case 3
'-----------------------------------------------------------------------------
    ASSET_COMPOUND_DAILY_RETURNS_FUNC = UPS_MATRIX
    'Identify UP days - UP = "1", DN = "0"
'-----------------------------------------------------------------------------
Case 4
'-----------------------------------------------------------------------------
    ASSET_COMPOUND_DAILY_RETURNS_FUNC = WEEKS_MATRIX
    'Identify the Day of the Week
'-----------------------------------------------------------------------------
Case Else
'-----------------------------------------------------------------------------
    ASSET_COMPOUND_DAILY_RETURNS_FUNC = Array(SUMMARY_MATRIX, DATA_MATRIX, DAILY_MATRIX, MONTHS_MATRIX, UPS_MATRIX, WEEKS_MATRIX)
'-----------------------------------------------------------------------------
End Select
'-----------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
ASSET_COMPOUND_DAILY_RETURNS_FUNC = Err.number
End Function
