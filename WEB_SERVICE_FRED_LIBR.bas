Attribute VB_Name = "WEB_SERVICE_FRED_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'---------------------------------------------------------------------------------
'---------------------------------------------------------------------------------
Private Type FRED_MAT_OBJ
    iSymbol As String
    iRefer As String
    iDate() As Variant
    iData() As Variant
End Type
'---------------------------------------------------------------------------------
'---------------------------------------------------------------------------------

'Function  grab economic data from the St. Louis Federal Reserve web site

Function FRED_GET_DATA_BY_DATE_FUNC(ByVal TICKER_STR As String, _
ByVal DATE_VAL As Date, _
Optional ByVal OUTPUT As Variant = 0)
'2011.04.28

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim TEMP1_STR As String
Dim TEMP2_STR As String
Dim TEMP_VAL As Variant
Dim DATA_STR As String
Dim DATE_STR As String
Dim TITLE_STR As String

Dim SRC_URL_STR As String

On Error GoTo ERROR_LABEL

If IsDate(DATE_VAL) = False Then: GoTo ERROR_LABEL

'See if web page has already been retrieved
SRC_URL_STR = "https://research.stlouisfed.org/fred2/data/" & UCase(TICKER_STR) & ".txt"
DATA_STR = SAVE_WEB_DATA_PAGE_FUNC(SRC_URL_STR, 0, False, 0, False)

TEMP2_STR = Chr(13)
FRED_GET_DATA_BY_DATE_FUNC = ""
Select Case True
Case UCase(OUTPUT) = "TITLE"
    TEMP1_STR = "Title:": GoSub MID_LINE
Case UCase(OUTPUT) = "SERIES ID"
    TEMP1_STR = "Series ID:": GoSub MID_LINE
Case UCase(OUTPUT) = "SOURCE"
    TEMP1_STR = "Source:": GoSub MID_LINE
Case UCase(OUTPUT) = "RELEASE"
    TEMP1_STR = "Release:": GoSub MID_LINE
Case UCase(OUTPUT) = "SEASONAL ADJUSTMENT"
    TEMP1_STR = "Seasonal Adjustment:": GoSub MID_LINE
Case UCase(OUTPUT) = "FREQUENCY"
    TEMP1_STR = "Frequency:": GoSub MID_LINE
Case UCase(OUTPUT) = "UNITS"
    TEMP1_STR = "Units:": GoSub MID_LINE
Case UCase(OUTPUT) = "DATE RANGE"
    TEMP1_STR = "Date Range:": GoSub MID_LINE
Case UCase(OUTPUT) = "LAST UPDATED"
    TEMP1_STR = "Last Updated:": GoSub MID_LINE
Case UCase(OUTPUT) = "NOTES"
    TEMP1_STR = "Notes:": TEMP2_STR = "DATE": GoSub MID_LINE
    h = 0
    Do While h <> Len(FRED_GET_DATA_BY_DATE_FUNC)
       h = Len(FRED_GET_DATA_BY_DATE_FUNC)
       FRED_GET_DATA_BY_DATE_FUNC = Replace(FRED_GET_DATA_BY_DATE_FUNC, "  ", " ")
    Loop
End Select
If FRED_GET_DATA_BY_DATE_FUNC <> "" Then Exit Function

1983:
'First Pass: Get title of file
i = InStr(DATA_STR, "Title:")
j = InStr(i, DATA_STR, Chr(13))
TITLE_STR = Trim(Mid(DATA_STR, i + 6, j - i - 6))

'Second Pass: Look for date in file
k = InStr(DATA_STR, "Notes:")

i = 0
l = CLng(DATE_VAL)
Do While i = 0
    DATE_STR = Format(DATE_VAL, "yyyy-mm-dd")
    i = InStr(k, DATA_STR, DATE_STR)
    DATE_VAL = DateAdd("d", -1, DATE_VAL)
    If l <= 0 Then: Exit Do
Loop
If i <= 0 Then GoTo ERROR_LABEL
   
j = InStr(i, DATA_STR, Chr(13))
TEMP_VAL = Mid(DATA_STR, i + 11, j - i - 11)
'On Error Resume Next
 '   TEMP_VAL = CDec(TEMP_VAL)
'On Error GoTo ERROR_LABEL
TEMP_VAL = CONVERT_STRING_NUMBER_FUNC(TEMP_VAL)

'---------------------------------------------------------------
Select Case OUTPUT
'---------------------------------------------------------------
Case 0
'---------------------------------------------------------------
    FRED_GET_DATA_BY_DATE_FUNC = TEMP_VAL
'---------------------------------------------------------------
Case Else
'---------------------------------------------------------------
    FRED_GET_DATA_BY_DATE_FUNC = Array(TEMP_VAL, DATE_STR, TITLE_STR)
'---------------------------------------------------------------
End Select
'---------------------------------------------------------------
    
Exit Function
'---------------------------------------------------------------
MID_LINE:
'---------------------------------------------------------------
    i = InStr(DATA_STR, TEMP1_STR)
    If i = 0 Then: GoTo 1983
    i = i + Len(TEMP1_STR)
    j = InStr(i, DATA_STR, TEMP2_STR)
    FRED_GET_DATA_BY_DATE_FUNC = Trim(Mid(DATA_STR, i, j - i))
'---------------------------------------------------------------
Return
'---------------------------------------------------------------
ERROR_LABEL:
FRED_GET_DATA_BY_DATE_FUNC = Err.number
End Function

       
'// PERFECT
                
Function FRED_GET_DATA_BY_SERIES_FUNC(ByVal TICKER_STR As String, _
Optional ByVal RESORT_FLAG As Boolean = False, _
Optional ByVal OUTPUT As Long = 0)

'FRED_GET_DATA_BY_SERIES_FUNC("CURRNS")

Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim SROW As Long
Dim NROWS As Long
Dim NSIZE As Long

Dim DATA_STR As String
Dim SRC_URL_STR As String

Dim TEMP_ARR As Variant
Dim DATA_ARR() As Variant

Dim DATA_VECTOR As Variant
Dim DATE_VECTOR As Variant

On Error GoTo ERROR_LABEL
        
GoSub SOURCE_LINE
'-------------------------------------------------------------------------------------
Select Case OUTPUT
'-------------------------------------------------------------------------------------
Case 0
'-------------------------------------------------------------------------------------
    ReDim DATA_VECTOR(1 To NSIZE, 1 To 2)
    j = 1: k = 10
    If RESORT_FLAG = False Then
        For i = SROW To NROWS: GoSub LOAD_LINE0: Next i
    Else
        For i = NROWS To SROW Step -1: GoSub LOAD_LINE0: Next i
    End If
    FRED_GET_DATA_BY_SERIES_FUNC = DATA_VECTOR
'-------------------------------------------------------------------------------------
Case Else
'-------------------------------------------------------------------------------------
    ReDim DATA_VECTOR(1 To NSIZE, 1 To 1)
    ReDim DATE_VECTOR(1 To NSIZE, 1 To 1)
    j = 1: k = 10
    If RESORT_FLAG = False Then
        For i = SROW To NROWS: GoSub LOAD_LINE1: Next i
    Else
        For i = NROWS To SROW Step -1: GoSub LOAD_LINE1: Next i
    End If
    FRED_GET_DATA_BY_SERIES_FUNC = Array(DATE_VECTOR, CONVERT_STRING_NUMBER_FUNC(DATA_VECTOR), DATA_STR)
'-------------------------------------------------------------------------------------
End Select
'-------------------------------------------------------------------------------------

'-------------------------------------------------------------------------------------
Exit Function
'-------------------------------------------------------------------------------------
SOURCE_LINE:
'-------------------------------------------------------------------------------------
    SRC_URL_STR = "http://research.stlouisfed.org/fred2/series/" & TICKER_STR & _
                  "/downloaddata/" & TICKER_STR & ".txt"
    DATA_STR = SAVE_WEB_DATA_PAGE_FUNC(SRC_URL_STR, 0, False, 0, False)
    DATA_STR = Replace(DATA_STR, Chr(13), "", 1, -1, vbBinaryCompare)
    TEMP_ARR = Split(DATA_STR, Chr(10), -1, vbBinaryCompare)
    SROW = LBound(TEMP_ARR): NROWS = UBound(TEMP_ARR)
       
    k = 11
    DATA_STR = ""
    For i = SROW To NROWS
        If i <= k Then: DATA_STR = DATA_STR & TEMP_ARR(i) & Chr(10)
        If (Left(TEMP_ARR(i), 4) = "DATE") And (Right(TEMP_ARR(i), 5) = "VALUE") Then
            j = i: Exit For
        End If
    Next i
    
    k = 1
    For i = j + 1 To NROWS
        If Trim(TEMP_ARR(i)) <> "" Then
            ReDim Preserve DATA_ARR(1 To k)
            DATA_ARR(k) = TEMP_ARR(i)
            k = k + 1
        End If
    Next i
    SROW = LBound(DATA_ARR): NROWS = UBound(DATA_ARR)
    NSIZE = NROWS - SROW + 1
'-------------------------------------------------------------------------------------
Return
'-------------------------------------------------------------------------------------
LOAD_LINE1:
'-------------------------------------------------------------------------------------
    TEMP_ARR = Mid(DATA_ARR(i), 1, k)
    TEMP_ARR = Split(TEMP_ARR, "-", -1, vbBinaryCompare)
    l = LBound(TEMP_ARR)
    DATE_VECTOR(j, 1) = DateSerial(CInt(TEMP_ARR(l)), CInt(TEMP_ARR(l + 1)), CInt(TEMP_ARR(l + 2)))
    DATA_VECTOR(j, 1) = Mid(DATA_ARR(i), k + 1, Len(DATA_ARR(i)))
    j = j + 1
'-------------------------------------------------------------------------------------
Return
'-------------------------------------------------------------------------------------
LOAD_LINE0:
'-------------------------------------------------------------------------------------
    TEMP_ARR = Mid(DATA_ARR(i), 1, k)
    TEMP_ARR = Split(TEMP_ARR, "-", -1, vbBinaryCompare)
    l = LBound(TEMP_ARR)
    DATA_VECTOR(j, 1) = DateSerial(CInt(TEMP_ARR(l)), CInt(TEMP_ARR(l + 1)), CInt(TEMP_ARR(l + 2)))
    DATA_VECTOR(j, 2) = CONVERT_STRING_NUMBER_FUNC(Mid(DATA_ARR(i), k + 1, Len(DATA_ARR(i))))
    j = j + 1
'-------------------------------------------------------------------------------------
Return
'-------------------------------------------------------------------------------------
ERROR_LABEL:
FRED_GET_DATA_BY_SERIES_FUNC = Err.number
End Function
           
'// PERFECT

Function FRED_DATA_SERIES_FUNC(ByRef TICKERS_RNG As Variant, _
Optional ByVal OUTPUT As Long = 3)

Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim NROWS As Long

Dim TICKER_STR As String

Dim TEMP_ARR As Variant
Dim DATA_GROUP() As FRED_MAT_OBJ
Dim TEMP_MATRIX As Variant
Dim TICKERS_VECTOR As Variant

On Error GoTo ERROR_LABEL

If IsArray(TICKERS_RNG) = True Then
    TICKERS_VECTOR = TICKERS_RNG
    If UBound(TICKERS_VECTOR, 1) = 1 Then: _
        TICKERS_VECTOR = MATRIX_TRANSPOSE_FUNC(TICKERS_VECTOR)
Else
    ReDim TICKERS_VECTOR(1 To 1, 1 To 1)
    TICKERS_VECTOR(1, 1) = TICKERS_RNG
End If

NROWS = UBound(TICKERS_VECTOR)

j = 0: k = 1
For i = 1 To NROWS
    TICKER_STR = TICKERS_VECTOR(i, 1)
    TEMP_ARR = FRED_GET_DATA_BY_SERIES_FUNC(TICKER_STR, True, 1)
    If IsArray(TEMP_ARR) = False Then: GoTo 1983
    l = UBound(TEMP_ARR(LBound(TEMP_ARR)), 1)
    If l >= k Then: k = l
    j = j + 1
    ReDim Preserve DATA_GROUP(1 To j)
    DATA_GROUP(j).iSymbol = TICKER_STR
    DATA_GROUP(j).iDate = TEMP_ARR(LBound(TEMP_ARR))
    DATA_GROUP(j).iData = TEMP_ARR(LBound(TEMP_ARR) + 1)
    DATA_GROUP(j).iRefer = TEMP_ARR(LBound(TEMP_ARR) + 2)
1983:
Next i

ReDim TEMP_MATRIX(1 To j, 1 To 3)
For i = 1 To j
    TEMP_MATRIX(i, 1) = DATA_GROUP(i).iSymbol
    TEMP_MATRIX(i, 2) = DATA_GROUP(i).iDate
    TEMP_MATRIX(i, 3) = DATA_GROUP(i).iData
Next i
Erase DATA_GROUP()
TEMP_MATRIX = SANITISED_DATES_VECTOR_FUNC(TEMP_MATRIX, "d", k, OUTPUT)
FRED_DATA_SERIES_FUNC = TEMP_MATRIX
    
Exit Function
ERROR_LABEL:
FRED_DATA_SERIES_FUNC = Err.number
End Function

Function FRED_TOP3_CATEGORIES_FUNC()
'2011.04.28

Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim TEMP_STR As String
Dim DATA_STR As String

Dim SRC_URL_STR As String
Dim TEMP_ARR() As String

On Error GoTo ERROR_LABEL

SRC_URL_STR = "http://research.stlouisfed.org/fred2/categories/"
DATA_STR = SAVE_WEB_DATA_PAGE_FUNC(SRC_URL_STR, 0, False, 0, False)
i = InStr(1, DATA_STR, ">Categories - Top 3 Levels</h1>")
If i = 0 Then: GoTo ERROR_LABEL

j = InStr(i, DATA_STR, "</table>")
If j = 0 Then: GoTo ERROR_LABEL
DATA_STR = Mid(DATA_STR, i, j - i)
DATA_STR = Replace(DATA_STR, "amp;", "")


i = InStr(1, DATA_STR, "href=")
If i = 0 Then: GoTo ERROR_LABEL
i = InStr(i, DATA_STR, Chr(34) & ">")
i = i + 2

j = InStr(i, DATA_STR, "<")
If j = 0 Then: GoTo ERROR_LABEL

l = 0
ReDim TEMP_ARR(1 To l + 1)
Do
    k = j
    TEMP_STR = Trim(Mid(DATA_STR, i, j - i))
    j = i - 2
    i = j
    Do
        i = i - 1
    Loop Until Mid(DATA_STR, i, 1) = "/"
    i = i + 1
    TEMP_STR = Trim(Mid(DATA_STR, i, j - i)) & "|" & TEMP_STR
    
    l = l + 1
    ReDim Preserve TEMP_ARR(1 To l)
    TEMP_ARR(l) = TEMP_STR
    
    i = InStr(k, DATA_STR, "href=")
    If i = 0 Then: Exit Do
    i = InStr(i, DATA_STR, Chr(34) & ">") + 2
    j = InStr(i, DATA_STR, "<")
Loop Until j = 0

ReDim TEMP_VECTOR(0 To l, 1 To 2)
TEMP_VECTOR(0, 1) = "INDEX"
TEMP_VECTOR(0, 2) = "CATEGORY"

For k = 1 To l
    TEMP_STR = TEMP_ARR(k)
    i = 1
    j = InStr(i, TEMP_STR, "|")
    TEMP_VECTOR(k, 1) = Mid(TEMP_STR, i, j - i)
    i = j + 1
    j = Len(TEMP_STR) + 1
    TEMP_VECTOR(k, 2) = Mid(TEMP_STR, i, j - i)
Next k

FRED_TOP3_CATEGORIES_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
FRED_TOP3_CATEGORIES_FUNC = Err.number
End Function

Function FRED_SERIES_CATEGORIES_FUNC(ByVal INDEX_CATEGORY As String)
'2011.04.28

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long
Dim m As Long
Dim n As Long

Dim TEMP_STR As String
Dim DATA_STR As String

Dim REFER_STR As String
Dim SERVER_STR As String
Dim SUFFIX_STR As String
Dim SRC_URL_STR As String
Dim TEMP_MATRIX() As String

On Error GoTo ERROR_LABEL

SERVER_STR = "http://research.stlouisfed.org/fred2/categories/"
SRC_URL_STR = SERVER_STR & INDEX_CATEGORY
DATA_STR = SAVE_WEB_DATA_PAGE_FUNC(SRC_URL_STR, 0, False, 0, False)

i = InStr(1, DATA_STR, "Series 1")
If i = 0 Then: GoTo ERROR_LABEL
i = InStr(i, DATA_STR, "of")
If i = 0 Then: GoTo ERROR_LABEL

i = i + 2
j = i
Do
    j = j + 1
Loop Until Mid(DATA_STR, j, 1) = "&"

n = Val(Trim(Mid(DATA_STR, i, j - i)))
If n = 0 Then: GoTo ERROR_LABEL

ReDim TEMP_MATRIX(0 To n, 1 To 8)

TEMP_MATRIX(0, 1) = ("SERIES ID")
TEMP_MATRIX(0, 2) = ("TITLE")
TEMP_MATRIX(0, 3) = ("START")
TEMP_MATRIX(0, 4) = ("END")
TEMP_MATRIX(0, 5) = ("FREQ")
TEMP_MATRIX(0, 6) = ("UNITS")
TEMP_MATRIX(0, 7) = ("SEAS ADJ")
TEMP_MATRIX(0, 8) = ("UPDATED")

h = 1
k = 1
l = 1
REFER_STR = Chr(34) & "checkbox" & Chr(34)
GoSub TRIM_LINE
Do While h <= n
    If k > 50 Then 'Trigger
        l = l + 1
        SUFFIX_STR = INDEX_CATEGORY & "?cid=" & INDEX_CATEGORY & "&pageID=" & l
        SRC_URL_STR = SERVER_STR & SUFFIX_STR
        DATA_STR = SAVE_WEB_DATA_PAGE_FUNC(SRC_URL_STR, 0, False, 0, False)
        GoSub TRIM_LINE
        k = 1
    End If
    i = InStr(1, DATA_STR, REFER_STR)
    If i = 0 Then: Exit Do
    Do While i <> 0
        For m = 1 To 2
            i = InStr(i, DATA_STR, Chr(34) & ">") + 2
        Next m
        For m = 1 To 8
            i = InStr(i, DATA_STR, Chr(34) & ">") + 2
            j = InStr(i, DATA_STR, "<")
            If j = 0 Then: GoTo 1983
            TEMP_STR = Mid(DATA_STR, i, j - i)
            TEMP_STR = Replace(TEMP_STR, Chr(10), "")
            TEMP_STR = REMOVE_EXTRA_SPACES_FUNC(TEMP_STR)
            TEMP_MATRIX(h, m) = Trim(TEMP_STR)
        Next m
1983:
        h = h + 1
        k = k + 1
        i = InStr(j, DATA_STR, REFER_STR)
    Loop
Loop

FRED_SERIES_CATEGORIES_FUNC = TEMP_MATRIX

Exit Function
'----------------------------------------------------------------
TRIM_LINE:
'----------------------------------------------------------------
i = InStr(1, DATA_STR, "Observation Range")
If i = 0 Then: GoTo ERROR_LABEL
j = Len(DATA_STR)
DATA_STR = Mid(DATA_STR, i, j - i)
Return
'----------------------------------------------------------------
ERROR_LABEL:
'----------------------------------------------------------------
FRED_SERIES_CATEGORIES_FUNC = Err.number
End Function
     
Public Function FRED_PRINT_HISTORICAL_DATA_FUNC( _
ByRef DST_RNG As Excel.Range, _
ByRef TICKERS_RNG As Variant, _
Optional ByVal VALID_FLAG As Boolean = False, _
Optional ByVal OUTPUT As Long = 0)

Dim SROW As Long
Dim NROWS As Long

Dim SCOLUMN As Long
Dim NCOLUMNS As Long

Dim RNG_ARR() As Excel.Range
Dim TEMP_FLAG As Boolean
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

FRED_PRINT_HISTORICAL_DATA_FUNC = False
TEMP_MATRIX = FRED_DATA_SERIES_FUNC(TICKERS_RNG, OUTPUT)

SROW = LBound(TEMP_MATRIX, 1)
NROWS = UBound(TEMP_MATRIX, 1)
SCOLUMN = LBound(TEMP_MATRIX, 2)
NCOLUMNS = UBound(TEMP_MATRIX, 2)

'For j = SCOLUMN To NCOLUMNS
 '   For i = SROW To NROWS
  '      DST_RNG.Cells(i, j) = TEMP_MATRIX(i, j)
   ' Next i
'Next j

Set DST_RNG = Range(DST_RNG.Cells(SROW, SCOLUMN), DST_RNG.Cells(NROWS, NCOLUMNS))
DST_RNG.value = TEMP_MATRIX

Select Case OUTPUT
Case 0 'Sanitised Matrix without Headings
    Call RNG_FORMAT_DATES_VECTOR_FUNC(DST_RNG, False, 1)
    Set DST_RNG = RNG_RESIZE_RNG_FUNC(DST_RNG, 0, 0)
Case 1 'Sanitised Matrix with Headings
    Call RNG_FORMAT_DATES_VECTOR_FUNC(DST_RNG, True, 1)
    Set DST_RNG = RNG_RESIZE_RNG_FUNC(DST_RNG.Offset(1, 0), 1, 0)

Case 2 'Sanitised Matrix & Match Dates without Headings
    Call RNG_FORMAT_DATES_VECTOR_FUNC(DST_RNG, False, 0)
    Set DST_RNG = RNG_RESIZE_RNG_FUNC(DST_RNG.Offset(0, 1), 0, 1)
Case Else 'Sanitised Matrix & Match Dates with Headings
    Call RNG_FORMAT_DATES_VECTOR_FUNC(DST_RNG, True, 0)
    If VALID_FLAG = True Then
        Excel.Application.ScreenUpdating = True
        TEMP_FLAG = RNG_FILL_SET_ARR_FUNC(DST_RNG, RNG_ARR())
        If TEMP_FLAG = False Then: GoTo 1983
        Call RNG_CHECK_BLANKS_FUNC(RNG_ARR(), 2, 1, 2, 1, 2)
    End If
1983:
    Set DST_RNG = RNG_RESIZE_RNG_FUNC(DST_RNG.Offset(1, 1), 1, 1)
End Select

FRED_PRINT_HISTORICAL_DATA_FUNC = True

Exit Function
ERROR_LABEL:
FRED_PRINT_HISTORICAL_DATA_FUNC = False
End Function

Function FRED_HISTORICAL_DATA_CHART_FUNC(ByVal TICKER_STR As String, _
Optional ByVal OUTPUT As Long = 0)
On Error GoTo ERROR_LABEL
Select Case OUTPUT
Case 0 'Quotes
    FRED_HISTORICAL_DATA_CHART_FUNC = _
    "http://research.stlouisfed.org/fred2/series/" & TICKER_STR & "/"
Case Else 'Chart
    FRED_HISTORICAL_DATA_CHART_FUNC = _
    "http://research.stlouisfed.org/fred2/data/" & TICKER_STR & "_Max_630_378.png"
End Select
Exit Function
ERROR_LABEL:
FRED_HISTORICAL_DATA_CHART_FUNC = Err.number
End Function
