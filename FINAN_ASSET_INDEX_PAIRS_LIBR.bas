Attribute VB_Name = "FINAN_ASSET_INDEX_PAIRS_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'Some people say that the DOW influences foreign markets ... usually.
'If the DOW is down today, chances are that foreign markets will be
'down tomorrow.

'So every morning (when the markets aren't having a holiday), I check
'the Asian and European market indexes. I'm hoping that I can tell
'whether North American markets will open Up or Down. Who influences whom?
'Of course, i 've already investigated the possibilities: DOW vs the World
'If foreign markets are Up (or Down) early one morning, will the DOW open Up
'(or Down) that day? If the DOW closes Up (or Down) one day, will foreign markets
'open Up (or Down) the next day? Alas, I never learn! I can't tell nothin' from
'foreign markets.

'The routine has been modified so you can choose to compare the returns of
'the Index to that of the second,third,.... N days later

'http://money.cnn.com/data/premarket/

Function INDEX_PAIRS_CORRELATION_FUNC(ByVal INDEX_STR As String, _
ByRef TICKERS_RNG As Variant, _
ByVal START_DATE As Date, _
ByVal END_DATE As Date, _
Optional ByVal PERIODS_LATER As Integer = 1, _
Optional ByVal OUTPUT As Integer = 0)

'You type in a couple of Yahoo stock symbols, pick an End Date then magic.
'the daily returns of the first with the daily returns of the second one day later.
'So you can see if the first influences the second, right? Something like that, and
'If you try it with a recent End Date and do it ag'in with an End Date a year ago,
'you often find a significant difference.

Dim i As Long
Dim j As Long
Dim k As Long
Dim NROWS As Long
Dim NCOLUMNS As Long
Dim TICKERS_VECTOR As Variant
Dim TEMP1_MATRIX As Variant
Dim TEMP2_MATRIX As Variant
Dim TEMP3_MATRIX As Variant

On Error GoTo ERROR_LABEL

If IsArray(TICKERS_RNG) = True Then
    TEMP1_MATRIX = TICKERS_RNG
    If UBound(TEMP1_MATRIX, 1) = 1 Then
        TEMP1_MATRIX = MATRIX_TRANSPOSE_FUNC(TEMP1_MATRIX)
    End If
    NCOLUMNS = UBound(TEMP1_MATRIX, 1)
    NCOLUMNS = NCOLUMNS + 1
    ReDim TICKERS_VECTOR(1 To NCOLUMNS, 1 To 1)
    TICKERS_VECTOR(1, 1) = INDEX_STR
    For i = 2 To NCOLUMNS
        TICKERS_VECTOR(i, 1) = TEMP1_MATRIX(i - 1, 1)
    Next i
    Erase TEMP1_MATRIX
Else
    ReDim TICKERS_VECTOR(1 To 2, 1 To 1)
    TICKERS_VECTOR(1, 1) = INDEX_STR
    TICKERS_VECTOR(2, 1) = TICKERS_RNG
    NCOLUMNS = 2
End If

TEMP1_MATRIX = YAHOO_HISTORICAL_DATA_SERIES1_FUNC(TICKERS_VECTOR, START_DATE, END_DATE, 6, "d", False, True)
TEMP2_MATRIX = TEMP1_MATRIX
TEMP2_MATRIX = MATRIX_REMOVE_COLUMNS_FUNC(TEMP2_MATRIX, 1, 1)
TEMP2_MATRIX = MATRIX_PERCENT_FUNC(TEMP2_MATRIX, 0)


'-------------------------------------------------------------------------
Select Case OUTPUT
'-------------------------------------------------------------------------
Case 0 'Correlation Matrix
'-------------------------------------------------------------------------
    NROWS = UBound(TEMP2_MATRIX, 1) - PERIODS_LATER
    ReDim TEMP1_MATRIX(1 To NROWS, 1 To NCOLUMNS)
    k = PERIODS_LATER
    For i = 1 To NROWS
        TEMP1_MATRIX(i, 1) = TEMP2_MATRIX(i, 1)
        For j = 2 To NCOLUMNS
            TEMP1_MATRIX(i, j) = TEMP2_MATRIX(k + 1, j)
        Next j
        k = k + 1
    Next i
    TEMP1_MATRIX = MATRIX_CORRELATION_FUNC(TEMP1_MATRIX)
    ReDim TEMP3_MATRIX(0 To NCOLUMNS, 0 To NCOLUMNS)
    TEMP3_MATRIX(0, 0) = "CORRELATION-MATRIX"
    For j = 1 To NCOLUMNS
        TEMP3_MATRIX(0, j) = TICKERS_VECTOR(j, 1)
        TEMP3_MATRIX(j, 0) = TEMP3_MATRIX(0, j)
        For i = j To NCOLUMNS
            TEMP3_MATRIX(i, j) = TEMP1_MATRIX(i, j)
        Next i
    Next j
'-------------------------------------------------------------------------
Case Else 'Returns
'-------------------------------------------------------------------------
    NCOLUMNS = NCOLUMNS + 1 'Dates Vector
    NROWS = UBound(TEMP2_MATRIX, 1) - PERIODS_LATER
    ReDim TEMP3_MATRIX(0 To NROWS, 1 To NCOLUMNS)
    TEMP3_MATRIX(0, 1) = "DATES"
    For j = 2 To NCOLUMNS
        TEMP3_MATRIX(0, j) = TICKERS_VECTOR(j - 1, 1)
    Next j
    k = PERIODS_LATER
    For i = 1 To NROWS
        TEMP3_MATRIX(i, 1) = TEMP1_MATRIX(i + 1, 1)
        TEMP3_MATRIX(i, 2) = TEMP2_MATRIX(i, 1)
        For j = 3 To NCOLUMNS
            TEMP3_MATRIX(i, j) = TEMP2_MATRIX(k + 1, j - 1)
        Next j
        k = k + 1
    Next i
'-------------------------------------------------------------------------
End Select
'-------------------------------------------------------------------------

Erase TEMP1_MATRIX
Erase TEMP2_MATRIX
Erase TICKERS_VECTOR
INDEX_PAIRS_CORRELATION_FUNC = TEMP3_MATRIX

Exit Function
ERROR_LABEL:
INDEX_PAIRS_CORRELATION_FUNC = Err.number
End Function



