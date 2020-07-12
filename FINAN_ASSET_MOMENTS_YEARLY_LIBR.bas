Attribute VB_Name = "FINAN_ASSET_MOMENTS_YEARLY_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

Function ASSET_YEAR_MONTH_GAIN_FUNC(ByVal TICKER_STR As Variant, _
Optional ByVal MONTH_INT As Integer = 5, _
Optional ByVal START_YEAR As Integer = 2006, _
Optional ByVal END_YEAR As Integer = 2008, _
Optional ByVal OUTPUT As Integer = 0)

Dim i As Long
Dim j As Long
Dim k As Long

Dim ii As Long
Dim jj As Long

Dim NROWS As Long
Dim NSIZE As Long

Dim START_DATE As Date
Dim END_DATE As Date

Dim TEMP_VECTOR As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

If MONTH_INT > 12 Or MONTH_INT < 1 Then: GoTo ERROR_LABEL
NSIZE = END_YEAR - START_YEAR + 1

If IsArray(TICKER_STR) = True Then
    DATA_MATRIX = TICKER_STR
Else
    START_DATE = DateSerial(START_YEAR, MONTH_INT, 1)
    START_DATE = EDATE_FUNC(START_DATE, -1)
    END_DATE = DateSerial(END_YEAR, MONTH_INT, 1)
    END_DATE = EDATE_FUNC(END_DATE, 12)
    DATA_MATRIX = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, START_DATE, END_DATE, "MONTHLY", "DA", False, False, True)
End If

NROWS = UBound(DATA_MATRIX, 1)

ReDim Preserve DATA_MATRIX(1 To NROWS, 1 To 4)

DATA_MATRIX(1, 3) = 0: DATA_MATRIX(1, 4) = 0

DATA_MATRIX(2, 3) = DATA_MATRIX(2, 2) / DATA_MATRIX(1, 2) - 1
DATA_MATRIX(2, 4) = 1 + DATA_MATRIX(2, 3)
For i = 3 To NROWS
    DATA_MATRIX(i, 3) = DATA_MATRIX(i, 2) / DATA_MATRIX(i - 1, 2) - 1
    DATA_MATRIX(i, 4) = DATA_MATRIX(i - 1, 4) * (1 + DATA_MATRIX(i, 3))
Next i

ReDim TEMP_VECTOR(0 To NSIZE, 1 To 5)

ii = 0: jj = 0

j = 1: k = 11
For i = 1 To NSIZE
    TEMP_VECTOR(i, 1) = DATA_MATRIX(j + 1, 1)
    TEMP_VECTOR(i, 2) = DATA_MATRIX(j + 1, 3)
    TEMP_VECTOR(i, 3) = ((DATA_MATRIX(j + 1 + k, 4) / DATA_MATRIX(j + 1, 4)) ^ (1 / k) - 1)
    
    If TEMP_VECTOR(i, 2) > TEMP_VECTOR(i, 3) Then
        TEMP_VECTOR(i, 4) = 1
        ii = ii + TEMP_VECTOR(i, 4)
    Else
        TEMP_VECTOR(i, 4) = 0
    End If
    
    If TEMP_VECTOR(i, 2) * TEMP_VECTOR(i, 3) > 0 Then
        TEMP_VECTOR(i, 5) = 1
        jj = jj + TEMP_VECTOR(i, 5)
    Else
        TEMP_VECTOR(i, 5) = 0
    End If
    
    j = j + 12
    If j + 1 + k > NROWS Then: Exit For
Next i

TEMP_VECTOR(0, 1) = "DATE"
TEMP_VECTOR(0, 2) = UCase(Format(DATA_MATRIX(2, 1), "mmmm")) & " GAIN"
TEMP_VECTOR(0, 3) = "REST OF THE YEAR"
TEMP_VECTOR(0, 4) = Format(ii / NSIZE, "0.00%")
TEMP_VECTOR(0, 5) = Format(jj / NSIZE, "0.00%") 'a positive or negative Month was
'a harbinger of a positive or negative year x% of the time

Select Case OUTPUT
Case 0
    ASSET_YEAR_MONTH_GAIN_FUNC = TEMP_VECTOR
Case 1
    ASSET_YEAR_MONTH_GAIN_FUNC = DATA_MATRIX
Case Else
    ASSET_YEAR_MONTH_GAIN_FUNC = Array(TEMP_VECTOR, DATA_MATRIX)
End Select

Exit Function
ERROR_LABEL:
ASSET_YEAR_MONTH_GAIN_FUNC = Err.number
End Function


'// PERFECT

Function ASSETS_ANNUAL_RETURNS_FUNC(ByRef TICKERS_RNG As Variant, _
ByVal START_DATE As Date, _
ByVal END_DATE As Date, _
Optional ByVal PERIODS_PER_YEAR As Long = 252)

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long
Dim NSIZE As Long
Dim NROWS As Long
Dim NCOLUMNS As Long

Dim RETURN_VAL As Double

Dim TICKERS_VECTOR As Variant
Dim DATA_MATRIX As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

If IsArray(TICKERS_RNG) = True Then
    TICKERS_VECTOR = TICKERS_RNG
    If UBound(TICKERS_VECTOR, 2) = 1 Then
        TICKERS_VECTOR = MATRIX_TRANSPOSE_FUNC(TICKERS_VECTOR)
    End If
Else
    ReDim TICKERS_VECTOR(1 To 1, 1 To 1)
    TICKERS_VECTOR(1, 1) = TICKERS_RNG
End If
NCOLUMNS = UBound(TICKERS_VECTOR, 2)

DATA_MATRIX = YAHOO_HISTORICAL_DATA_SERIES1_FUNC(TICKERS_RNG, START_DATE, END_DATE, 6, "d", False, True) 'Ascending
NROWS = UBound(DATA_MATRIX, 1)
NSIZE = Int(NROWS / PERIODS_PER_YEAR)
If NROWS Mod PERIODS_PER_YEAR = 0 Then h = 0 Else h = 1
ReDim TEMP_MATRIX(0 To NSIZE + h, 1 To NCOLUMNS + 2)
TEMP_MATRIX(0, 1) = "STARTING PERIOD"
TEMP_MATRIX(0, 2) = "ENDING PERIOD"
For j = 1 To NCOLUMNS: TEMP_MATRIX(0, 2 + j) = TICKERS_VECTOR(1, j): Next j

j = 1
For k = 1 To NSIZE
    i = j: j = i + PERIODS_PER_YEAR
    TEMP_MATRIX(k, 1) = DATA_MATRIX(i, 1)
    TEMP_MATRIX(k, 2) = DATA_MATRIX(j, 1)
    GoSub CALC_VAL
Next k
If h = 1 Then
    k = NSIZE + 1
    i = j: j = NROWS
    TEMP_MATRIX(k, 1) = DATA_MATRIX(i, 1)
    TEMP_MATRIX(k, 2) = DATA_MATRIX(j, 1)
    GoSub CALC_VAL
End If

ASSETS_ANNUAL_RETURNS_FUNC = TEMP_MATRIX

'-----------------------------------------------------------------------------------------------------------------------
Exit Function
'-----------------------------------------------------------------------------------------------------------------------
CALC_VAL:
'-----------------------------------------------------------------------------------------------------------------------
    For l = 1 To NCOLUMNS
        If DATA_MATRIX(i, l + 1) <> 0 Then
            RETURN_VAL = DATA_MATRIX(j, l + 1) / DATA_MATRIX(i, l + 1) - 1
            TEMP_MATRIX(k, l + 2) = RETURN_VAL
        Else
            TEMP_MATRIX(k, l + 2) = ""
        End If
    Next l
'-----------------------------------------------------------------------------------------------------------------------
Return
'-----------------------------------------------------------------------------------------------------------------------
ERROR_LABEL:
ASSETS_ANNUAL_RETURNS_FUNC = Err.number
End Function


