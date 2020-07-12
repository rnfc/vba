Attribute VB_Name = "FINAN_ASSET_PROCESS_MATCH_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'Find the "Best" No Periods - day Match

'Over the past month (ending Jan 18, 2008) the markets have been dropping dramatically:
'the DOW down about a thousand points.
'I was beginning to wonder when these indices would reach bottom, so I
'thought I'd do the following:
'1. Assume the future is a replica of the past.
'2. Look carefully at the DOW performance during the past month.
'3. Examine every period in the last eight years.
'4. Identify that period which matches the current period.
'5. Determine how long that historical period took to recover.
'6. Predict that the current market chaos will end as it did in the
'   period succeeding the historical image.

'Having adopted 1 and completed 2, 3 and 4 ... I identified the best-match
'historical period

Function ASSET_MATCH_BEST_HIST_PERIODS_PREDICT1_FUNC(ByRef TICKER_STR As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal END_DATE As Date, _
Optional ByVal BACK_TEST_PERIODS As Long = 25, _
Optional ByVal FORWARD_PERIODS As Long = 0, _
Optional ByVal ERROR_TYPE As Integer = 1, _
Optional ByVal VERSION_PERIODS As Integer = 0, _
Optional ByVal FREQUENCY As Integer = 0, _
Optional ByRef HOLIDAYS_RNG As Variant)

'History never repeats itself but it does rhyme: "Mark Twain".
'BACK_TEST_PERIODS: Periods Back from the ending period

'0  Minimize Difference in Final Values
'1  Minimize Average Error
'2  Minimize Maximum Error
'>2 Minimize (1 - Correlation)

Dim h As Long
Dim i As Long 'Periods Back
Dim j As Long
Dim k As Long

Dim NROWS As Long
Dim NSIZE As Long

Dim MAX_VAL As Double
Dim TEMP_SUM As Double
Dim ERROR_VAL As Double
Dim XTEMP_VAL As Double
Dim YTEMP_VAL As Double

Dim PERIOD_STR As String
Dim TEMP1_VECTOR As Variant
Dim TEMP2_VECTOR As Variant

Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

If IsArray(TICKER_STR) = True Then
    DATA_MATRIX = TICKER_STR
Else
    If FREQUENCY = 0 Then: PERIOD_STR = "d"
    If FREQUENCY = 1 Then: PERIOD_STR = "w"
    If FREQUENCY >= 2 Then: PERIOD_STR = "m"
    DATA_MATRIX = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, START_DATE, END_DATE, PERIOD_STR, "DA", False, False, True)
End If
NROWS = UBound(DATA_MATRIX, 1)

NSIZE = BACK_TEST_PERIODS + FORWARD_PERIODS
ReDim TEMP_MATRIX(0 To NSIZE, 1 To 8)

TEMP_MATRIX(0, 1) = "CURRENT DATE"
TEMP_MATRIX(0, 2) = "CURRENT PRICE"
TEMP_MATRIX(0, 3) = "CURRENT GROWTH"
TEMP_MATRIX(0, 4) = "PAST DATE"
TEMP_MATRIX(0, 5) = "PAST PRICE"
TEMP_MATRIX(0, 6) = "PAST/FUTURE GROWTH"
TEMP_MATRIX(0, 8) = "PREDICTED PRICES"

'----------------------------------------------------------------------
If VERSION_PERIODS = 0 Then
'----------------------------------------------------------------------
    ERROR_VAL = 2 ^ 52
    For i = BACK_TEST_PERIODS To NROWS - BACK_TEST_PERIODS
        GoSub ERROR_LINE
        If YTEMP_VAL < ERROR_VAL Then
            j = i
            ERROR_VAL = YTEMP_VAL
        End If
    Next i
    i = j
    GoSub ERROR_LINE
    ERROR_VAL = YTEMP_VAL
'----------------------------------------------------------------------
Else
'----------------------------------------------------------------------
    If VERSION_PERIODS < BACK_TEST_PERIODS Then: VERSION_PERIODS = BACK_TEST_PERIODS
    i = VERSION_PERIODS
    GoSub ERROR_LINE
    j = i
    ERROR_VAL = YTEMP_VAL
'----------------------------------------------------------------------
End If
'----------------------------------------------------------------------
For k = 1 To BACK_TEST_PERIODS
    h = NROWS - BACK_TEST_PERIODS + k
    TEMP_MATRIX(k, 1) = DATA_MATRIX(h, 1)
    TEMP_MATRIX(k, 2) = DATA_MATRIX(h, 2)
    TEMP_MATRIX(k, 8) = "" 'TEMP_MATRIX(k, 2)
    TEMP_MATRIX(k, 3) = TEMP_MATRIX(k, 2) / TEMP_MATRIX(1, 2) - 1
    TEMP_MATRIX(k, 4) = DATA_MATRIX(h - i, 1)
    TEMP_MATRIX(k, 5) = DATA_MATRIX(h - i, 2)
    TEMP_MATRIX(k, 6) = TEMP_MATRIX(k, 5) / TEMP_MATRIX(1, 5) - 1
    TEMP_MATRIX(k, 7) = Abs(TEMP_MATRIX(k, 6) - TEMP_MATRIX(k, 3))
Next k
'----------------------------------------------------------------------
For k = 1 To FORWARD_PERIODS
    h = BACK_TEST_PERIODS + k
    Select Case FREQUENCY
    Case 0 'Daily
        TEMP_MATRIX(h, 1) = WORKDAY2_FUNC(TEMP_MATRIX(h - 1, 1), 1, HOLIDAYS_RNG)
    Case 1 'Weekly
        TEMP_MATRIX(h, 1) = WORKDAY2_FUNC(TEMP_MATRIX(h - 1, 1), 5, HOLIDAYS_RNG)
    Case Else 'Monthly
        TEMP_MATRIX(h, 1) = EDATE_FUNC(TEMP_MATRIX(h - 1, 1), 1)
    End Select
    TEMP_MATRIX(h, 2) = ""
    TEMP_MATRIX(h, 3) = ""
    
    TEMP_MATRIX(h, 4) = DATA_MATRIX(NROWS + k - j, 1)
    TEMP_MATRIX(h, 5) = DATA_MATRIX(NROWS + k - j, 2)
    TEMP_MATRIX(h, 6) = TEMP_MATRIX(h, 5) / TEMP_MATRIX(1, 5) - 1
    TEMP_MATRIX(h, 7) = ""
    If k <> 1 Then
        TEMP_MATRIX(h, 8) = TEMP_MATRIX(h - 1, 8) * TEMP_MATRIX(h, 5) / TEMP_MATRIX(h - 1, 5)
    Else
        TEMP_MATRIX(h, 8) = TEMP_MATRIX(h - 1, 2) * TEMP_MATRIX(h, 5) / TEMP_MATRIX(h - 1, 5)
    End If
Next k
TEMP_MATRIX(0, 7) = "MIN ERROR : " & Format(ERROR_VAL, "0.00%")

ASSET_MATCH_BEST_HIST_PERIODS_PREDICT1_FUNC = TEMP_MATRIX

'----------------------------------------------------------------------
Exit Function
'----------------------------------------------------------------------------
ERROR_LINE:
'----------------------------------------------------------------------------
    TEMP_SUM = 0
    MAX_VAL = -2 ^ 52
    For k = 1 To BACK_TEST_PERIODS
        h = NROWS - BACK_TEST_PERIODS + k
        XTEMP_VAL = (DATA_MATRIX(h - i, 2) / DATA_MATRIX(NROWS - BACK_TEST_PERIODS + 1 - i, 2) - 1)
        XTEMP_VAL = Abs(XTEMP_VAL - (DATA_MATRIX(h, 2) / DATA_MATRIX(NROWS - BACK_TEST_PERIODS + 1, 2) - 1))
        If XTEMP_VAL > MAX_VAL Then: MAX_VAL = XTEMP_VAL
        TEMP_SUM = TEMP_SUM + XTEMP_VAL
    Next k
    YTEMP_VAL = 0
    '----------------------------------------------------------------------------
    Select Case ERROR_TYPE
    '----------------------------------------------------------------------------
    Case 0 'Minimize Difference in Final Values
        YTEMP_VAL = XTEMP_VAL
    Case 1 'Minimize Average Error
        YTEMP_VAL = TEMP_SUM / BACK_TEST_PERIODS
    Case 2 'Minimize Maximum Error
        YTEMP_VAL = MAX_VAL
    Case Else 'Minimize (1 - Correlation)
        ReDim TEMP1_VECTOR(1 To BACK_TEST_PERIODS, 1 To 1)
        ReDim TEMP2_VECTOR(1 To BACK_TEST_PERIODS, 1 To 1)
        For k = 1 To BACK_TEST_PERIODS
            h = NROWS - BACK_TEST_PERIODS + k
            TEMP1_VECTOR(k, 1) = DATA_MATRIX(h, 2)
            TEMP2_VECTOR(k, 1) = DATA_MATRIX(h - i, 2)
        Next k
        YTEMP_VAL = 1 - CORRELATION_FUNC(TEMP1_VECTOR, TEMP2_VECTOR, 0, 0)
        Erase TEMP1_VECTOR
        Erase TEMP2_VECTOR
    End Select
'----------------------------------------------------------------------------
Return
'----------------------------------------------------------------------------

ERROR_LABEL:
ASSET_MATCH_BEST_HIST_PERIODS_PREDICT1_FUNC = Err.number
End Function

'We predicted a TSX of 13,818.84 by April 7, 2008, up 4.0%. It actually closed at 13,745, up 3.5%
'We predicted a DOW of 12,033.31 by April 7, 2008, up 1.2%. It actually closed at 12,612, up 6.0%
'We predicted MSFT at 27.50 by April 7, 2008, down 1.3%. It actually closed at 29.16, up 6.7%

'However, I was chatting on FWF about some other topic and Shakes mentioned auto correlation.
'That got me to thinking that, instead of minimizing the maximum error (between the last month and
'historical months), maybe I should maximize the correlation between them. That is, I start with a
'$1K portfolio and calculate the correlation between the subsequent portfolios and ...

Function ASSET_MATCH_BEST_HIST_PERIODS_PREDICT2_FUNC(ByRef DATE_RNG As Variant, _
ByRef DATA_RNG As Variant, _
Optional ByVal BACK_TEST_PERIODS As Long = 25, _
Optional ByVal FORWARD_PERIODS As Long = 18, _
Optional ByVal ERROR_TYPE As Integer = 3, _
Optional ByVal FREQUENCY As Integer = 0, _
Optional ByRef HOLIDAYS_RNG As Variant)

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim NROWS As Long
Dim NSIZE As Long

Dim MIN_VAL As Double
Dim MAX_VAL As Double
Dim TEMP_SUM As Double
Dim ERROR_VAL As Double

Dim TEMP1_VECTOR As Variant
Dim TEMP2_VECTOR As Variant

Dim DATA_VECTOR As Variant
Dim DATE_VECTOR As Variant

Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATE_VECTOR = DATE_RNG
If UBound(DATE_VECTOR, 1) = 1 Then
    DATE_VECTOR = MATRIX_TRANSPOSE_FUNC(DATE_VECTOR)
End If

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then
    DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
End If
If UBound(DATE_VECTOR, 1) <> UBound(DATA_VECTOR, 1) Then: GoTo ERROR_LABEL
NROWS = UBound(DATA_VECTOR, 1)

GoSub ERROR_LINE

NSIZE = BACK_TEST_PERIODS + FORWARD_PERIODS
ReDim TEMP_MATRIX(0 To NSIZE, 1 To 7)
TEMP_MATRIX(0, 1) = "CURRENT DATE"
TEMP_MATRIX(0, 2) = "CURRENT PRICE"
TEMP_MATRIX(0, 3) = "CURRENT $1 GROWTH"
TEMP_MATRIX(0, 4) = "PAST DATE"
TEMP_MATRIX(0, 5) = "PAST PRICE"
TEMP_MATRIX(0, 6) = "PAST $1 GROWTH"
TEMP_MATRIX(0, 7) = "MIN ERROR: " & Format(MIN_VAL, "0.00%")

h = 0
For j = BACK_TEST_PERIODS To 1 Step -1
    TEMP_MATRIX(j, 1) = DATE_VECTOR(i - h, 1)
    TEMP_MATRIX(j, 2) = DATA_VECTOR(i - h, 1)
    TEMP_MATRIX(j, 3) = DATA_VECTOR(i - h, 1) / DATA_VECTOR(i - BACK_TEST_PERIODS + 1, 1) 'Growth
    
    TEMP_MATRIX(j, 4) = DATE_VECTOR(l - h, 1)
    TEMP_MATRIX(j, 5) = DATA_VECTOR(l - h, 1)
    TEMP_MATRIX(j, 6) = DATA_VECTOR(l - h, 1) / DATA_VECTOR(l - BACK_TEST_PERIODS + 1, 1) 'Growth
    TEMP_MATRIX(j, 7) = Abs(TEMP_MATRIX(j, 6) - TEMP_MATRIX(j, 3))
    
    h = h + 1
Next j
h = 1
For j = BACK_TEST_PERIODS + 1 To NSIZE
    Select Case FREQUENCY
    Case 0 'Daily
        TEMP_MATRIX(j, 1) = WORKDAY2_FUNC(TEMP_MATRIX(j - 1, 1), 1, HOLIDAYS_RNG)
    Case 1 'Weekly
        TEMP_MATRIX(j, 1) = WORKDAY2_FUNC(TEMP_MATRIX(j - 1, 1), 5, HOLIDAYS_RNG)
    Case Else 'Monthly
        TEMP_MATRIX(j, 1) = EDATE_FUNC(TEMP_MATRIX(j - 1, 1), 1)
    End Select
    TEMP_MATRIX(j, 3) = ""
    TEMP_MATRIX(j, 4) = DATE_VECTOR(l + h, 1)
    TEMP_MATRIX(j, 5) = DATA_VECTOR(l + h, 1)
    TEMP_MATRIX(j, 6) = TEMP_MATRIX(j, 5) / TEMP_MATRIX(1, 5)
    TEMP_MATRIX(j, 7) = ""
    
    TEMP_MATRIX(j, 2) = TEMP_MATRIX(j - 1, 2) * TEMP_MATRIX(j, 6) / TEMP_MATRIX(j - 1, 6)
    
    h = h + 1
Next j
ASSET_MATCH_BEST_HIST_PERIODS_PREDICT2_FUNC = TEMP_MATRIX

'----------------------------------------------------------------------------------------------------
Exit Function
'----------------------------------------------------------------------------------------------------
ERROR_LINE:
'----------------------------------------------------------------------------------------------------
    MIN_VAL = 2 ^ 52
    i = NROWS
    Select Case ERROR_TYPE
    '----------------------------------------------------------------------------------------------------
    Case 0 'Minimize Difference in Final Values
    '----------------------------------------------------------------------------------------------------
        For j = i - BACK_TEST_PERIODS To BACK_TEST_PERIODS Step -1
            h = 0
            For k = j To j '- BACK_TEST_PERIODS + 1 Step -1
                ERROR_VAL = DATA_VECTOR(k, 1) / DATA_VECTOR(j - BACK_TEST_PERIODS + 1, 1)
                ERROR_VAL = Abs(ERROR_VAL - DATA_VECTOR(i - h, 1) / DATA_VECTOR(i - BACK_TEST_PERIODS + 1, 1)) '^ 2
                h = h + 1
            Next k
            If ERROR_VAL < MIN_VAL Then
                MIN_VAL = ERROR_VAL
                l = j
            End If
        Next j
    '----------------------------------------------------------------------------------------------------
    Case 1 'Minimize Average Error
    '----------------------------------------------------------------------------------------------------
        For j = i - BACK_TEST_PERIODS To BACK_TEST_PERIODS Step -1
            h = 0
            TEMP_SUM = 0
            For k = j To j - BACK_TEST_PERIODS + 1 Step -1
                ERROR_VAL = DATA_VECTOR(k, 1) / DATA_VECTOR(j - BACK_TEST_PERIODS + 1, 1)
                TEMP_SUM = TEMP_SUM + Abs(ERROR_VAL - DATA_VECTOR(i - h, 1) / DATA_VECTOR(i - BACK_TEST_PERIODS + 1, 1)) '^ 2
                h = h + 1
            Next k
            ERROR_VAL = (TEMP_SUM / BACK_TEST_PERIODS) '^ 0.5
            If ERROR_VAL < MIN_VAL Then
                MIN_VAL = ERROR_VAL
                l = j
            End If
        Next j
    '----------------------------------------------------------------------------------------------------
    Case 2 'Minimize Maximum Error
    '----------------------------------------------------------------------------------------------------
        For j = i - BACK_TEST_PERIODS To BACK_TEST_PERIODS Step -1
            h = 0
            MAX_VAL = -2 ^ 52
            For k = j To j - BACK_TEST_PERIODS + 1 Step -1
                ERROR_VAL = DATA_VECTOR(k, 1) / DATA_VECTOR(j - BACK_TEST_PERIODS + 1, 1)
                ERROR_VAL = Abs(ERROR_VAL - DATA_VECTOR(i - h, 1) / DATA_VECTOR(i - BACK_TEST_PERIODS + 1, 1)) '^ 2
                If ERROR_VAL > MAX_VAL Then: MAX_VAL = ERROR_VAL
                h = h + 1
            Next k
            ERROR_VAL = MAX_VAL
            If ERROR_VAL < MIN_VAL Then
                MIN_VAL = ERROR_VAL
                l = j
            End If
        Next j
    '----------------------------------------------------------------------------------------------------
    Case Else 'Minimize (1 - Correlation)
    '----------------------------------------------------------------------------------------------------
        ReDim TEMP1_VECTOR(1 To BACK_TEST_PERIODS, 1 To 1)
        ReDim TEMP2_VECTOR(1 To BACK_TEST_PERIODS, 1 To 1)
        For j = i - BACK_TEST_PERIODS To BACK_TEST_PERIODS Step -1
            h = 0
            For k = j To j - BACK_TEST_PERIODS + 1 Step -1
                TEMP1_VECTOR(BACK_TEST_PERIODS - h, 1) = DATA_VECTOR(i - h, 1) / DATA_VECTOR(i - BACK_TEST_PERIODS + 1, 1)
                TEMP2_VECTOR(BACK_TEST_PERIODS - h, 1) = DATA_VECTOR(k, 1)
                h = h + 1
            Next k
            ERROR_VAL = 1 - CORRELATION_FUNC(TEMP1_VECTOR, TEMP2_VECTOR, 0, 0)
            If ERROR_VAL < MIN_VAL Then
                MIN_VAL = ERROR_VAL
                l = j
            End If
        Next j
        Erase TEMP1_VECTOR
        Erase TEMP2_VECTOR
    '----------------------------------------------------------------------------------------------------
    End Select
'--------------------------------------------------------------------------------------------------------
Return
'--------------------------------------------------------------------------------------------------------
ERROR_LABEL:
'--------------------------------------------------------------------------------------------------------
ASSET_MATCH_BEST_HIST_PERIODS_PREDICT2_FUNC = Err.number
End Function

Function ASSET_MATCH_BEST_HIST_PERIODS_PREDICT3_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal SCOLUMN As Long = 7, _
Optional ByVal NO_LAGS As Long = 0, _
Optional ByVal PERIODS_BACKWARD As Long = 25, _
Optional ByVal PERIODS_FORWARD As Long = 18, _
Optional ByVal ERROR_TYPE As Integer = 3)
'http://www.gummy-stuff.org/historical-compare2.htm

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long
Dim m As Long
Dim n As Long
Dim o As Long
Dim NROWS As Long

Dim DIFF_VAL As Double
Dim ERROR_VAL As Double
Dim FACTOR_VAL As Double

Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant
Dim TEMP1_VECTOR As Variant
Dim TEMP2_VECTOR As Variant

On Error GoTo ERROR_LABEL

FACTOR_VAL = 1
DATA_MATRIX = DATA_RNG
NROWS = UBound(DATA_MATRIX, 1)
NROWS = NROWS - NO_LAGS
h = NROWS - PERIODS_BACKWARD
m = h + 1
n = 1
ReDim TEMP_MATRIX(1 To h - PERIODS_BACKWARD + 1, 1 To 3 + 1 + PERIODS_FORWARD) 'Index/Date/Error

For i = h To PERIODS_BACKWARD Step -1 'Find the "Best" z-day Match
    l = m
    k = i - PERIODS_BACKWARD + 1
    ERROR_VAL = 0
    '---------------------------------------------------------------------------------------------------------------
    Select Case ERROR_TYPE
    '---------------------------------------------------------------------------------------------------------------
    Case 0 'Minimize Difference in Final Values
    '---------------------------------------------------------------------------------------------------------------
            ERROR_VAL = Abs(FACTOR_VAL * DATA_MATRIX(i, SCOLUMN) / DATA_MATRIX(k, SCOLUMN) - _
                            FACTOR_VAL * DATA_MATRIX(NROWS, SCOLUMN) / DATA_MATRIX(m, SCOLUMN))
    '---------------------------------------------------------------------------------------------------------------
    Case 1 'Minimize Average Error
    '---------------------------------------------------------------------------------------------------------------
        For j = k To i
            DIFF_VAL = Abs(FACTOR_VAL * DATA_MATRIX(j, SCOLUMN) / DATA_MATRIX(k, SCOLUMN) - _
                           FACTOR_VAL * DATA_MATRIX(l, SCOLUMN) / DATA_MATRIX(m, SCOLUMN))
            ERROR_VAL = ERROR_VAL + DIFF_VAL
            l = l + 1
        Next j
        ERROR_VAL = ERROR_VAL / PERIODS_BACKWARD
    '---------------------------------------------------------------------------------------------------------------
    Case 2 'Minimize Maximum Error
    '---------------------------------------------------------------------------------------------------------------
        For j = k To i
            DIFF_VAL = Abs(FACTOR_VAL * DATA_MATRIX(j, SCOLUMN) / DATA_MATRIX(k, SCOLUMN) - _
                           FACTOR_VAL * DATA_MATRIX(l, SCOLUMN) / DATA_MATRIX(m, SCOLUMN))
            If DIFF_VAL > ERROR_VAL Then: ERROR_VAL = DIFF_VAL
            l = l + 1
        Next j
    '---------------------------------------------------------------------------------------------------------------
    Case Else 'Minimize (1 - Correlation)
    '---------------------------------------------------------------------------------------------------------------
        ReDim TEMP1_VECTOR(1 To i - k + 1, 1 To 1)
        ReDim TEMP2_VECTOR(1 To i - k + 1, 1 To 1)
        o = 1
        For j = k To i
            TEMP1_VECTOR(o, 1) = FACTOR_VAL * DATA_MATRIX(j, SCOLUMN) / DATA_MATRIX(k, SCOLUMN)
            TEMP2_VECTOR(o, 1) = FACTOR_VAL * DATA_MATRIX(l, SCOLUMN) / DATA_MATRIX(m, SCOLUMN)
            o = o + 1
            l = l + 1
        Next j
        ERROR_VAL = 1 - CORRELATION_FUNC(TEMP1_VECTOR, TEMP2_VECTOR, 0, 0)
    '---------------------------------------------------------------------------------------------------------------
    End Select
    '---------------------------------------------------------------------------------------------------------------
    TEMP_MATRIX(n, 1) = ERROR_VAL
    TEMP_MATRIX(n, 2) = DATA_MATRIX(k, 1)
    TEMP_MATRIX(n, 3) = DATA_MATRIX(i, 1)
    'Forecasting!!!!!!!!
    j = 1
    TEMP_MATRIX(n, 3 + 0 + j) = DATA_MATRIX(NROWS, SCOLUMN) 'Current Price
    TEMP_MATRIX(n, 3 + 1 + j) = DATA_MATRIX(NROWS, SCOLUMN) * DATA_MATRIX(i + j, SCOLUMN) / DATA_MATRIX(i + j - 1, SCOLUMN) 'Growth
    For j = 2 To PERIODS_FORWARD
        TEMP_MATRIX(n, 3 + 1 + j) = TEMP_MATRIX(n, 3 + j) * DATA_MATRIX(i + j, SCOLUMN) / DATA_MATRIX(i + j - 1, SCOLUMN) 'Growth
    Next j
    n = n + 1
Next i

TEMP_MATRIX = MATRIX_QUICK_SORT_FUNC(TEMP_MATRIX, 1, 1)
'Min Error/Starting Period/Ending Period/Pt0,Pt+1,Pt+2,...Ptt+n
ASSET_MATCH_BEST_HIST_PERIODS_PREDICT3_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
ASSET_MATCH_BEST_HIST_PERIODS_PREDICT3_FUNC = Err.number
End Function
