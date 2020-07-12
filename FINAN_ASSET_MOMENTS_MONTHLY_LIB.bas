Attribute VB_Name = "FINAN_ASSET_MOMENTS_MONTHLY_LIB"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

Function ASSET_MONTHLY_TREND_ALERT_FUNC(ByRef TICKER_STR As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal END_DATE As Date, _
Optional ByVal PERCENT1_VAL As Double = 0.01, _
Optional ByVal PERCENT2_VAL As Double = 0.1, _
Optional ByRef REFERENCE_RNG As Variant)

'Number Days: lows in previous x days

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim l() As Long

Dim NROWS As Long
Dim DATE_VAL As Date
Dim DATA_MATRIX As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

If IsArray(TICKER_STR) = True Then
    DATA_MATRIX = TICKER_STR
Else
    DATA_MATRIX = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, START_DATE, END_DATE, "M", "DA", False, True, True)
End If
NROWS = UBound(DATA_MATRIX, 1)

ReDim TEMP_MATRIX(0 To NROWS, 1 To 8)
TEMP_MATRIX(0, 1) = "DATE"
TEMP_MATRIX(0, 2) = "PRICE"
TEMP_MATRIX(0, 3) = "CHANGE"
TEMP_MATRIX(0, 4) = "PERCENT"
TEMP_MATRIX(0, 5) = "TREND" 'The trend is "Up" if the close is up more that 1 percent from
'the previous close, "Down" if the close is down more than 1 percent and "Flat" otherwise.
TEMP_MATRIX(0, 6) = "ALERTS" 'The high volatility alert is displayed if the change from the
'previous month is greater than 8 percent.
TEMP_MATRIX(0, 7) = "ABSOLUTE CHANGE"
TEMP_MATRIX(0, 8) = "ABSOLUTE PERCENT"

i = 1
TEMP_MATRIX(i, 1) = DATA_MATRIX(i, 1)
TEMP_MATRIX(i, 2) = DATA_MATRIX(i, 2)
For j = 3 To 8: TEMP_MATRIX(i, j) = "": Next j
For i = 2 To NROWS
    TEMP_MATRIX(i, 1) = DATA_MATRIX(i, 1)
    TEMP_MATRIX(i, 2) = DATA_MATRIX(i, 2)
    TEMP_MATRIX(i, 3) = TEMP_MATRIX(i, 2) - TEMP_MATRIX(i - 1, 2) 'Delta Price
    TEMP_MATRIX(i, 4) = TEMP_MATRIX(i, 3) / TEMP_MATRIX(i - 1, 2) 'Monthly Growth
    If TEMP_MATRIX(i, 4) > PERCENT1_VAL Then
        TEMP_MATRIX(i, 5) = "Up"
    Else
        If TEMP_MATRIX(i, 4) < -PERCENT1_VAL Then
            TEMP_MATRIX(i, 5) = "Down"
        Else
            TEMP_MATRIX(i, 5) = "Flat"
        End If
    End If
    If TEMP_MATRIX(i, 4) >= PERCENT2_VAL Then
        TEMP_MATRIX(i, 6) = "High Volatility"
    Else
        If TEMP_MATRIX(i, 4) <= -PERCENT2_VAL Then
            TEMP_MATRIX(i, 6) = "High Volatility"
        Else
            TEMP_MATRIX(i, 6) = ""
        End If
    End If
    TEMP_MATRIX(i, 7) = Abs(TEMP_MATRIX(i, 3))
    TEMP_MATRIX(i, 8) = Abs(TEMP_MATRIX(i, 4))
Next i

If IsArray(REFERENCE_RNG) = False Then
    ASSET_MONTHLY_TREND_ALERT_FUNC = TEMP_MATRIX
Else
    Dim NSIZE As Long
    Dim INDEX_OBJ As Collection
    Dim TEMP_VECTOR As Variant
    Dim REFERENCE_VECTOR As Variant
    REFERENCE_VECTOR = REFERENCE_RNG
    If UBound(REFERENCE_VECTOR, 1) = 1 Then
        REFERENCE_VECTOR = MATRIX_TRANSPOSE_FUNC(REFERENCE_VECTOR)
    End If
    
    NSIZE = UBound(REFERENCE_VECTOR, 1)
    ReDim TEMP_VECTOR(0 To NSIZE, 1 To 8)
    TEMP_VECTOR(0, 1) = "STARTING PERIOD"
    TEMP_VECTOR(0, 2) = "ENDING PERIOD"
    TEMP_VECTOR(0, 3) = "AVERAGE CHANGE"
    TEMP_VECTOR(0, 4) = "AVERAGE PERCENT"
    TEMP_VECTOR(0, 5) = "PREDOMINANT TREND"
    TEMP_VECTOR(0, 6) = "NUMBER OF ALERTS"
    TEMP_VECTOR(0, 7) = "MAX CHANGE"
    TEMP_VECTOR(0, 8) = "MAX PERCENT"
    
    On Error Resume Next
    Set INDEX_OBJ = New Collection
    For k = 1 To NROWS
        DATE_VAL = DateSerial(Year(DATA_MATRIX(k, 1)), Month(DATA_MATRIX(k, 1)), 1)
        Call INDEX_OBJ.Add(CStr(k), CStr(DATE_VAL))
        If Err.number <> 0 Then: Err.Clear
    Next k
    
    For k = 1 To NSIZE
        DATE_VAL = DateSerial(Year(REFERENCE_VECTOR(k, 1)), Month(REFERENCE_VECTOR(k, 1)), 1)
        i = Val(INDEX_OBJ.Item(CStr(DATE_VAL)))
        TEMP_VECTOR(k, 1) = DATE_VAL
        If i = 0 Or Err.number <> 0 Then
            Err.Clear
            GoTo 1983
        End If
        DATE_VAL = DateSerial(Year(REFERENCE_VECTOR(k, 2)), Month(REFERENCE_VECTOR(k, 2)), 1)
        j = Val(INDEX_OBJ.Item(CStr(DATE_VAL)))
        TEMP_VECTOR(k, 2) = DATE_VAL
        If j = 0 Or Err.number <> 0 Then
            Err.Clear
            GoTo 1983
        End If
        
        TEMP_VECTOR(k, 3) = 0: TEMP_VECTOR(k, 4) = 0
        TEMP_VECTOR(k, 7) = -2 ^ 52: TEMP_VECTOR(k, 8) = -2 ^ 52
        ReDim l(1 To 4)
        For h = i To j
            TEMP_VECTOR(k, 3) = TEMP_VECTOR(k, 3) + TEMP_MATRIX(h, 3)
            TEMP_VECTOR(k, 4) = TEMP_VECTOR(k, 4) + TEMP_MATRIX(h, 4)
            If TEMP_MATRIX(h, 7) > TEMP_VECTOR(k, 7) Then: TEMP_VECTOR(k, 7) = TEMP_MATRIX(h, 7)
            If TEMP_MATRIX(h, 8) > TEMP_VECTOR(k, 8) Then: TEMP_VECTOR(k, 8) = TEMP_MATRIX(h, 8)
            Select Case TEMP_MATRIX(h, 5)
            Case "Up"
                l(1) = l(1) + 1
            Case "Down"
                l(2) = l(2) + 1
            Case "Flat"
                l(3) = l(3) + 1
            End Select
            If TEMP_MATRIX(h, 6) = "High Volatility" Then: l(4) = l(4) + 1
        Next h
        TEMP_VECTOR(k, 3) = TEMP_VECTOR(k, 3) / (j - i + 1)
        TEMP_VECTOR(k, 4) = TEMP_VECTOR(k, 4) / (j - i + 1)
        TEMP_VECTOR(k, 6) = l(4)
        If l(2) > l(1) Then
            TEMP_VECTOR(k, 5) = "Down"
        Else
            If l(1) > l(2) Then
                TEMP_VECTOR(k, 5) = "Up"
            Else
                TEMP_VECTOR(k, 5) = "Flat"
            End If
        End If
1983:
    Next k
    Erase TEMP_MATRIX: Erase DATA_MATRIX: Set INDEX_OBJ = Nothing
    ASSET_MONTHLY_TREND_ALERT_FUNC = TEMP_VECTOR
End If


Exit Function
ERROR_LABEL:
ASSET_MONTHLY_TREND_ALERT_FUNC = Err.number
End Function


Function ASSETS_MONTHLY_ROC_FUNC(ByRef TICKERS_RNG As Variant, _
ByVal START_DATE As Date, _
ByVal END_DATE As Date)

Dim i As Long
Dim j As Long
Dim k As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim MEAN_VAL As Double
Dim SIGMA_VAL As Double
Dim RETURN_VAL As Double

Dim P_ARR() As Double
Dim N_ARR() As Double

Dim TICKER_STR As String
Dim TICKERS_VECTOR As Variant
Dim DATA_VECTOR As Variant
Dim TEMP_MATRIX As Variant

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

ReDim TEMP_MATRIX(1 To 17, 1 To NCOLUMNS + 1)
TEMP_MATRIX(1, 1) = "TICKER"
TEMP_MATRIX(2, 1) = "STARTING PERIOD"
TEMP_MATRIX(3, 1) = "ENDING PERIOD"
TEMP_MATRIX(4, 1) = "VOLATILITY"
TEMP_MATRIX(5, 1) = "AVERAGE RETURN"

i = 1
For k = 6 To 17
    TEMP_MATRIX(k, 1) = UCase(Format(DateSerial(Year(Date), i, Day(Date)), "mmm"))
    i = i + 1
Next k

For k = 1 To NCOLUMNS
    TICKER_STR = TICKERS_VECTOR(k, 1)
    TEMP_MATRIX(1, k + 1) = TICKER_STR
    DATA_VECTOR = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, START_DATE, END_DATE, "MONTHLY", "DA", False, False, True)
    If IsArray(DATA_VECTOR) = False Then: GoTo 1983
    NROWS = UBound(DATA_VECTOR, 1)
    If NROWS <= 1 Then: GoTo 1983

    TEMP_MATRIX(2, k + 1) = DATA_VECTOR(1, 1)
    TEMP_MATRIX(3, k + 1) = DATA_VECTOR(NROWS, 1)

    ReDim P_ARR(1 To 12)
    ReDim N_ARR(1 To 12)
    
    MEAN_VAL = 0
    For i = 2 To NROWS
        If IsDate(DATA_VECTOR(i, 1)) = False Then: GoTo 1983
        j = Month(DATA_VECTOR(i, 1))
        If DATA_VECTOR(i - 1, 2) <> 0 Then
            RETURN_VAL = DATA_VECTOR(i, 2) / DATA_VECTOR(i - 1, 2) - 1
        Else
            RETURN_VAL = 0
        End If
        P_ARR(j) = P_ARR(j) + RETURN_VAL
        N_ARR(j) = N_ARR(j) + 1
        MEAN_VAL = MEAN_VAL + RETURN_VAL
    Next i
    MEAN_VAL = MEAN_VAL / (NROWS - 1)
    
    SIGMA_VAL = 0
    For i = 2 To NROWS
        If DATA_VECTOR(i - 1, 2) <> 0 Then
            RETURN_VAL = DATA_VECTOR(i, 2) / DATA_VECTOR(i - 1, 2) - 1
        Else
            RETURN_VAL = 0
        End If
        SIGMA_VAL = SIGMA_VAL + (MEAN_VAL - RETURN_VAL) ^ 2
    Next i
    
    SIGMA_VAL = (SIGMA_VAL / (NROWS - 1)) ^ 0.5
    For j = 1 To 12
        If N_ARR(j) <> 0 Then
            TEMP_MATRIX(j + 5, k + 1) = P_ARR(j) / N_ARR(j)
        Else
            TEMP_MATRIX(j + 5, k + 1) = 0
        End If
    Next j

    TEMP_MATRIX(4, k + 1) = SIGMA_VAL
    TEMP_MATRIX(5, k + 1) = MEAN_VAL

1983:
Next k

ASSETS_MONTHLY_ROC_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
ASSETS_MONTHLY_ROC_FUNC = Err.number
End Function

'With few exceptions, stocks have larger gains on Halloween than on other end-of-month days.
'It seems there's something called the Halloween Indicator.
'It suggests you be in Cash during the Summer months (starting in May) then Buy after
'Halloween (Nov 1). Isn 't that "Sell in May then Go Away"?
'Yes! Have you heard of it? No.
'Well, we decided to test the idea and got this, comparing the total (cumulative) gain fo the period
'(Nov-to-Apr) and that for the period (May-to-Oct). We looked at the 10-years from Jan 1, 1999 to
'Jan 1, 2009 ... and got this for various country indexes....
'I see it's getting close to the end of October. Excuse me while I Buy some stocks ...

'http://en.wikipedia.org/wiki/Halloween_indicator

Function ASSETS_HALLOWEEN_INDICATOR_FUNC(ByRef TICKERS_RNG As Variant, _
ByVal START_DATE As Date, _
ByVal END_DATE As Date, _
Optional ByVal START_MONTH As Integer = 11, _
Optional ByVal END_MONTH As Integer = 4)

'--------------------------------------------------------------------------------------------
'Symbols Labels
'--------------------------------------------------------------------------------------------
'^AORD   Australia
'^BVSP   Brazil
'^GSPTSE Canada
'000001.SS   China
'^DJI    Dow
'^CAC    France
'^GDAXI  Germany
'^HSI    Hong Kong
'^BSESN  India
'^TA100  Israel
'^N225   Japan
'^KLSE   Malaysia
'^MXX    Mexico
'^IXIC   Nasdaq
'^GSPC   S&P500
'^KS11   S. Korea
'^STI    Singapore
'^TWII   Taiwan
'--------------------------------------------------------------------------------------------

Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long
Dim NROWS As Long
Dim NCOLUMNS As Long

Dim END1_STR As String
Dim END2_STR As String

Dim START1_STR As String
Dim START2_STR As String

Dim TICKER_STR As String

Dim RETURN_VAL As Double
Dim TICKERS_VECTOR As Variant
Dim DATA_VECTOR As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

'-------------------------------------------------------------------------------------------------------------------------
If IsArray(TICKERS_RNG) = True Then
'-------------------------------------------------------------------------------------------------------------------------
    TICKERS_VECTOR = TICKERS_RNG
    If UBound(TICKERS_VECTOR, 1) = 1 Then
        TICKERS_VECTOR = MATRIX_TRANSPOSE_FUNC(TICKERS_VECTOR)
    End If

    NCOLUMNS = UBound(TICKERS_VECTOR, 1)
    
    k = Year(Now)
    END1_STR = UCase(Format(DateSerial(k, END_MONTH, 1), "mmm"))
    START1_STR = UCase(Format(DateSerial(k, START_MONTH, 1), "mmm"))
    
    START2_STR = UCase(Format(DateSerial(k, IIf((END_MONTH + 1) > 12, 1, END_MONTH + 1), 1), "mmm"))
    END2_STR = UCase(Format(DateSerial(k, IIf((START_MONTH - 1) < 1, 12, START_MONTH - 1), 1), "mmm"))
    
    ReDim TEMP_MATRIX(0 To NCOLUMNS, 1 To 6)
    TEMP_MATRIX(0, 1) = "TICKER"
    TEMP_MATRIX(0, 2) = "STARTING PERIOD"
    TEMP_MATRIX(0, 3) = "ENDING PERIOD"
    TEMP_MATRIX(0, 4) = "GROWTH " & START2_STR & " - " & END2_STR
    TEMP_MATRIX(0, 5) = "GROWTH " & START1_STR & " - " & END1_STR
    TEMP_MATRIX(0, 6) = "( " & START1_STR & " - " & END1_STR & " ) - ( " & START2_STR & " - " & END2_STR & " )"
    
    For j = 1 To NCOLUMNS
        TICKER_STR = TICKERS_VECTOR(j, 1)
        
        TEMP_MATRIX(j, 1) = TICKER_STR
        DATA_VECTOR = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, START_DATE, END_DATE, "DAILY", "DA", False, False, True)

        If IsArray(DATA_VECTOR) = False Then: GoTo 1983
        NROWS = UBound(DATA_VECTOR, 1)
        If NROWS <= 1 Then: GoTo 1983
    
        TEMP_MATRIX(j, 2) = DATA_VECTOR(1, 1)
        TEMP_MATRIX(j, 3) = DATA_VECTOR(NROWS, 1)
        TEMP_MATRIX(j, 4) = 1: TEMP_MATRIX(j, 5) = 1
        For i = 2 To NROWS
            RETURN_VAL = DATA_VECTOR(i, 2) / DATA_VECTOR(i - 1, 2)
            l = Month(DATA_VECTOR(i, 1))
            If (l > END_MONTH And l < START_MONTH) Then
                TEMP_MATRIX(j, 4) = TEMP_MATRIX(j, 4) * RETURN_VAL
                TEMP_MATRIX(j, 5) = TEMP_MATRIX(j, 5) * 1
            Else
                TEMP_MATRIX(j, 4) = TEMP_MATRIX(j, 4) * 1
                TEMP_MATRIX(j, 5) = TEMP_MATRIX(j, 5) * RETURN_VAL
            End If
        Next i
        TEMP_MATRIX(j, 4) = TEMP_MATRIX(j, 4) - 1
        TEMP_MATRIX(j, 5) = TEMP_MATRIX(j, 5) - 1
        TEMP_MATRIX(j, 6) = TEMP_MATRIX(j, 5) - TEMP_MATRIX(j, 4)
1983:
    Next j
'-------------------------------------------------------------------------------------------------------------------------
Else
'Comparing the average return on the last day of October to the average of all
'end-of-month returns. I look at the difference between these two averages and
'see that Hobgoblins actually like halloween.
'-------------------------------------------------------------------------------------------------------------------------
    ReDim TEMP_MATRIX(0 To 13, 1 To 3)
    TICKER_STR = UCase(TICKERS_RNG)
    TEMP_MATRIX(0, 1) = TICKER_STR
    TEMP_MATRIX(0, 2) = "MEAN"
    TEMP_MATRIX(0, 3) = "VOLATILITY"
    
    TEMP_MATRIX(1, 1) = "ALL"
    i = 1
    For j = 2 To 13
        TEMP_MATRIX(j, 1) = UCase(Format(DateSerial(Year(Date), i, Day(Date)), "mmm"))
        i = i + 1
    Next j
        
    DATA_VECTOR = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, START_DATE, END_DATE, "DAILY", "DA", False, False, True)
    If IsArray(DATA_VECTOR) = False Then: GoTo 1983
    NROWS = UBound(DATA_VECTOR, 1)
    If NROWS <= 1 Then: GoTo ERROR_LABEL
    ReDim TICKERS_VECTOR(1 To 12)
    For j = 1 To 12: TICKERS_VECTOR(j) = 0: Next j
    For i = 2 To NROWS - 1
        l = Month(DATA_VECTOR(i, 1))
        If l <> Month(DATA_VECTOR(i + 1, 1)) Then
            RETURN_VAL = DATA_VECTOR(i, 2) / DATA_VECTOR(i - 1, 2) - 1
            TEMP_MATRIX(1, 2) = TEMP_MATRIX(1, 2) + RETURN_VAL
            k = k + 1
            j = l
            TEMP_MATRIX(j + 1, 2) = TEMP_MATRIX(j + 1, 2) + RETURN_VAL
            TICKERS_VECTOR(j) = TICKERS_VECTOR(j) + 1
        End If
    Next i
    For j = 1 To 12: TEMP_MATRIX(j + 1, 2) = TEMP_MATRIX(j + 1, 2) / TICKERS_VECTOR(j): Next j
    TEMP_MATRIX(1, 2) = TEMP_MATRIX(1, 2) / k
    For i = 2 To NROWS - 1
        l = Month(DATA_VECTOR(i, 1))
        If l <> Month(DATA_VECTOR(i + 1, 1)) Then
            RETURN_VAL = DATA_VECTOR(i, 2) / DATA_VECTOR(i - 1, 2) - 1
            TEMP_MATRIX(1, 3) = TEMP_MATRIX(1, 3) + (TEMP_MATRIX(1, 2) - RETURN_VAL) ^ 2
            j = l
            TEMP_MATRIX(j + 1, 3) = TEMP_MATRIX(j + 1, 3) + (TEMP_MATRIX(j + 1, 2) - RETURN_VAL) ^ 2
        End If
    Next i
    For j = 1 To 12: TEMP_MATRIX(j + 1, 3) = (TEMP_MATRIX(j + 1, 3) / TICKERS_VECTOR(j)) ^ 0.5: Next j
    TEMP_MATRIX(1, 3) = (TEMP_MATRIX(1, 3) / k) ^ 0.5
'-------------------------------------------------------------------------------------------------------------------------
End If
'-------------------------------------------------------------------------------------------------------------------------

ASSETS_HALLOWEEN_INDICATOR_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
ASSETS_HALLOWEEN_INDICATOR_FUNC = Err.number
End Function


'What percentage of (annualized) 12-month returns were greater than X?
'If I were to invest in stock XYZ for 24 months, what are my chances of getting a
'CAGR greater than 10%?

'Example: Trimark Mutual Fund, daily values, from Jan 1/97 to Nov 26/99).
'You select the time period (say, 26 days) and it looks at ALL 26 day periods (from Jan 1/97
'to Nov 26/99) and computes the percentage of times you would have made a Gain (over these 26
'days) of over 10% or over 20% or less than -5% etc.

'Select 250 days and you get (roughly) all one-year returns, ending on ANY day in the period
'Jan 6/98 to Nov 26/99. Uh ... I really should point out that 250 "market" days from Jan 1/97
'is Jan 6/98. P.S. ATP means All Time Periods, a technique championed by Wilfred Vos.

'http://www.gummy-stuff.org/ATP.htm
'http://www.gummy-stuff.org/sorta_ATP.htm

Function ASSETS_MONTHLY_ATP_FUNC(ByRef TICKERS_RNG As Variant, _
ByVal START_DATE As Date, _
ByVal END_DATE As Date, _
Optional ByVal MIN_RETURN As Double = -0.01, _
Optional ByVal DELTA_RETURN As Double = 0.01, _
Optional ByVal NBINS As Long = 8)

Dim i As Long
Dim j As Long
Dim k As Long
Dim h As Long

Dim NROWS As Long
Dim NCOLUMNS As Long
Dim TICKER_STR As String
Dim RETURN_VAL As Double
Dim DATA_VECTOR As Variant
Dim TICKERS_VECTOR As Variant

On Error GoTo ERROR_LABEL

h = 12
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

ReDim TEMP_MATRIX(0 To NCOLUMNS, 1 To NBINS + 1)
TEMP_MATRIX(0, 1) = "SYMBOL/RETURN"

RETURN_VAL = MIN_RETURN
For j = 1 To NBINS
    TEMP_MATRIX(0, j + 1) = RETURN_VAL
    RETURN_VAL = RETURN_VAL + DELTA_RETURN
Next j

For k = 1 To NCOLUMNS
    TICKER_STR = TICKERS_VECTOR(k, 1)
    TEMP_MATRIX(k, 1) = TICKER_STR
    DATA_VECTOR = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, START_DATE, END_DATE, "MONTHLY", "A", False, False, True)
    If IsArray(DATA_VECTOR) = False Then: GoTo 1983
    NROWS = UBound(DATA_VECTOR, 1)
    For i = 2 To NROWS - h + 1
        RETURN_VAL = 1
        For j = i To i + h - 1: RETURN_VAL = RETURN_VAL * DATA_VECTOR(j, 1) / DATA_VECTOR(j - 1, 1): Next j
        RETURN_VAL = RETURN_VAL ^ (12 / h) - 1
        For j = 1 To NBINS
            If RETURN_VAL > TEMP_MATRIX(0, j + 1) Then: TEMP_MATRIX(k, j + 1) = TEMP_MATRIX(k, j + 1) + 1
        Next j
    Next i
    For j = 1 To NBINS: TEMP_MATRIX(k, j + 1) = TEMP_MATRIX(k, j + 1) / (NROWS - h + 1): Next j
1983:
Next k

ASSETS_MONTHLY_ATP_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
ASSETS_MONTHLY_ATP_FUNC = Err.number
End Function
