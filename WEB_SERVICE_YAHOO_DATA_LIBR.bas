Attribute VB_Name = "WEB_SERVICE_YAHOO_DATA_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

Public Type YAHOO_DATA_OBJ 'Written by Nico for the Ben Graham Centre for Value Investing
    iSymbol As String
    iData() As Variant
End Type

Private PUB_YAHOO_HASH_OBJ As clsTypeHash

Function YAHOO_HISTORICAL_DATA_SERIE_FUNC(ByVal TICKER_STR As String, _
ByVal START_DATE As Date, _
ByVal END_DATE As Date, _
Optional ByVal PERIOD_STR As String = "MONTHLY", _
Optional ByRef ELEMENT_STR As String = "DOHLCVA", _
Optional ByVal HEADERS_FLAG As Boolean = False, _
Optional ByVal ADJUST_FLAG As Boolean = False, _
Optional ByVal RESORT_FLAG As Boolean = True)

On Error GoTo ERROR_LABEL

If UCase(PERIOD_STR) = "MONTHLY" Then PERIOD_STR = "m" 'Chooser for Monthly Data
If UCase(PERIOD_STR) = "DAILY" Then PERIOD_STR = "d" 'Chooser for Daily Data
If UCase(PERIOD_STR) = "WEEKLY" Then PERIOD_STR = "w" 'Chooser for Weekly Data

YAHOO_HISTORICAL_DATA_SERIE_FUNC = YAHOO_HISTORICAL_DATA_MATRIX_FUNC(TICKER_STR, Year(START_DATE), Month(START_DATE), Day(START_DATE), Year(END_DATE), Month(END_DATE), Day(END_DATE), PERIOD_STR, ELEMENT_STR, HEADERS_FLAG, ADJUST_FLAG, RESORT_FLAG, 0, 0)

Exit Function
ERROR_LABEL:
YAHOO_HISTORICAL_DATA_SERIE_FUNC = Err.number
End Function

Function YAHOO_HISTORICAL_DATA_SERIES1_FUNC(ByRef TICKERS_RNG As Variant, _
ByVal START_DATE As Variant, _
ByVal END_DATE As Variant, _
Optional ByVal ELEMENT_INT As Integer = 6, _
Optional ByRef PERIOD_STR As String = "d", _
Optional ByVal HEADERS_FLAG As Boolean = False, _
Optional ByVal RESORT_FLAG As Boolean = True)

'ELEMENT_INT (1)Open; (2)High; (3)Low; (4)Close; (5) Volume; (6)Adj.Close
'IF RESORT_FLAG = True Then: Ascending Order else Descending

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim hh As Long
Dim ii As Long
Dim jj As Long
Dim kk As Long

Dim NROWS As Long
Dim NCOLUMNS As Long
Dim KEY_STR As String
Dim CHR_STR As String
Dim DELIM_STR As String

Dim DATE_VAL As String
Dim DATA_VAL As String
Dim DATA_STR As String
Dim TICKER_STR As String
Dim SRC_URL_STR As String

Dim TEMP_STR As String
Dim LEFT_STR As String
Dim RIGHT_STR As String
Dim LINE_STR As String
Dim ERROR_STR As String

Dim END_DAY_INT As Long
Dim END_YEAR_INT As Long
Dim END_MONTH_INT As Long

Dim START_DAY_INT As Long
Dim START_YEAR_INT As Long
Dim START_MONTH_INT As Long

Dim DATE_VECTOR As Variant
Dim TEMP_MATRIX As Variant
Dim TICKERS_VECTOR As Variant

Dim DATA_OBJ As clsTypeHash

On Error GoTo ERROR_LABEL

If PUB_YAHOO_HASH_OBJ Is Nothing Then
    Set PUB_YAHOO_HASH_OBJ = New clsTypeHash
    PUB_YAHOO_HASH_OBJ.SetSize 10000
    PUB_YAHOO_HASH_OBJ.IgnoreCase = False
End If

CHR_STR = Chr(10)
DELIM_STR = ","
PERIOD_STR = LCase(PERIOD_STR)
If ELEMENT_INT < 1 Then: ELEMENT_INT = 1
If ELEMENT_INT > 6 Then: ELEMENT_INT = 6

START_YEAR_INT = Year(START_DATE)
START_MONTH_INT = Month(START_DATE)
START_DAY_INT = Day(START_DATE)
    
END_YEAR_INT = Year(END_DATE)
END_MONTH_INT = Month(END_DATE)
END_DAY_INT = Day(END_DATE)
   
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

KEY_STR = "": For jj = 1 To NCOLUMNS: KEY_STR = KEY_STR & "|" & TICKERS_VECTOR(jj, 1): Next jj
KEY_STR = KEY_STR & "|" & CStr(START_YEAR_INT) & "|" & CStr(START_MONTH_INT) & "|" & CStr(START_DAY_INT) & "|" & _
                          CStr(END_YEAR_INT) & "|" & "|" & CStr(END_MONTH_INT) & "|" & CStr(END_DAY_INT) & "|" & _
                          CStr(ELEMENT_INT) & "|" & PERIOD_STR & "|" & CStr(HEADERS_FLAG) & "|" & CStr(RESORT_FLAG)

If PUB_YAHOO_HASH_OBJ.Exists(KEY_STR) Then
    TEMP_MATRIX = PUB_YAHOO_HASH_OBJ(KEY_STR)
    YAHOO_HISTORICAL_DATA_SERIES1_FUNC = CONVERT_STRING_NUMBER_FUNC(TEMP_MATRIX)
    Exit Function
End If

Set DATA_OBJ = New clsTypeHash
DATA_OBJ.SetSize 10000
DATA_OBJ.IgnoreCase = False

LINE_STR = DELIM_STR
For jj = 1 To NCOLUMNS: LINE_STR = LINE_STR & "0" & DELIM_STR: Next jj
hh = 0
For jj = 1 To NCOLUMNS
    TICKER_STR = TICKERS_VECTOR(jj, 1)
    SRC_URL_STR = YAHOO_HISTORICAL_DATA_URL_FUNC(TICKER_STR, START_YEAR_INT, _
    START_MONTH_INT, START_DAY_INT, END_YEAR_INT, END_MONTH_INT, _
    END_DAY_INT, PERIOD_STR, NROWS, ERROR_STR)
    
    If ERROR_STR <> "" Then: GoTo 1983
    DATA_STR = SAVE_WEB_DATA_PAGE_FUNC(SRC_URL_STR, 0, False, 0, False)
    If DATA_STR = PUB_WEB_DATA_SYSTEM_ERROR_STR Then: GoTo 1983
    NROWS = COUNT_CHARACTERS_FUNC(DATA_STR, CHR_STR) - 1 'Number of periods
    DATA_STR = Replace(DATA_STR, CHR_STR, DELIM_STR & CHR_STR)
    
    h = 1
    j = InStr(h, DATA_STR, CHR_STR) 'Skip headings
    If j = 0 Then: GoTo 1983
    h = j + 1
    
    For ii = 1 To NROWS
        j = InStr(h, DATA_STR, DELIM_STR)
        If j = 0 Then Exit For
        DATE_VAL = Mid(DATA_STR, h, j - h) 'yyyy-mm-dd
        DATE_VAL = CStr(NORMALIZE_DATES_VECTOR_FUNC(CDate(DATE_VAL), PERIOD_STR))
                
        If DATA_OBJ.Exists(DATE_VAL) = False Then
            TEMP_STR = DATE_VAL & LINE_STR
            Call DATA_OBJ.Add(DATE_VAL, TEMP_STR)
            hh = hh + 1
        End If
        
        i = j + 1
        DATA_VAL = "0"
        For kk = 1 To ELEMENT_INT 'o/h/l/c/v/adj.c
            j = InStr(i, DATA_STR, DELIM_STR)
            If j = 0 Then GoTo 1982
            DATA_VAL = Mid(DATA_STR, i, j - i)
            i = j + 1
        Next kk

        If DATA_VAL <> "0" Then
            TEMP_STR = DATA_OBJ.Item(DATE_VAL)
            
            i = 1
            j = InStr(i, TEMP_STR, DELIM_STR)
            If j = 0 Then GoTo 1982
            i = j + 1
            For kk = 1 To jj
                LEFT_STR = Mid(TEMP_STR, 1, i - 1)
                j = InStr(i, TEMP_STR, DELIM_STR)
                If j = 0 Then GoTo 1982
                RIGHT_STR = Mid(TEMP_STR, j, Len(TEMP_STR) - j + 1)
                i = j + 1
            Next kk
            
            DATA_OBJ.Item(DATE_VAL) = LEFT_STR & DATA_VAL & RIGHT_STR
        End If
1982:
        h = InStr(h, DATA_STR, CHR_STR)
        If h = 0 Then: Exit For
        h = h + 1
    Next ii

1983:
Next jj

ReDim DATE_VECTOR(1 To hh, 1 To 1)
For ii = 1 To hh
    h = ii - 1
    DATE_VAL = DATA_OBJ.GetKey(h)
    DATE_VECTOR(ii, 1) = CDate(DATE_VAL)
Next ii
DATE_VECTOR = MATRIX_QUICK_SORT_FUNC(DATE_VECTOR, 1, IIf(RESORT_FLAG = True, 1, 0))

'-------------------------------------------------------------------------------------------------
If HEADERS_FLAG = True Then
'-------------------------------------------------------------------------------------------------
    l = 0
    ReDim TEMP_MATRIX(l To hh, 1 To NCOLUMNS + 1)
    TEMP_MATRIX(l, 1) = "DATES"
    For jj = 1 To NCOLUMNS: TEMP_MATRIX(l, jj + 1) = TICKERS_VECTOR(jj, 1): Next jj
'-------------------------------------------------------------------------------------------------
Else
'-------------------------------------------------------------------------------------------------
    l = 1: ReDim TEMP_MATRIX(l To hh, 1 To NCOLUMNS + 1)
'-------------------------------------------------------------------------------------------------
End If
'-------------------------------------------------------------------------------------------------

For ii = 1 To hh
    TEMP_MATRIX(ii, 1) = DATE_VECTOR(ii, 1)
    DATE_VAL = DATE_VECTOR(ii, 1)
    DATA_VAL = DATA_OBJ.Item(DATE_VAL)
    k = Len(DATA_VAL)
    i = InStr(1, DATA_VAL, DELIM_STR)
    If i = 0 Then: GoTo 1984
    i = i + 1
    jj = 1
    Do
        j = InStr(i, DATA_VAL, DELIM_STR)
        If j = 0 Then: Exit Do
        TEMP_MATRIX(ii, jj + 1) = CDec(Mid(DATA_VAL, i, j - i))
        i = j + 1
        jj = jj + 1
    Loop Until jj > NCOLUMNS
1984:
Next ii

Set DATA_OBJ = Nothing
Call PUB_YAHOO_HASH_OBJ.Add(KEY_STR, TEMP_MATRIX)
YAHOO_HISTORICAL_DATA_SERIES1_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
YAHOO_HISTORICAL_DATA_SERIES1_FUNC = Err.number
End Function

Function YAHOO_HISTORICAL_DATA_SERIES2_FUNC(ByRef TICKERS_RNG As Variant, _
ByVal START_DATE As Variant, _
ByVal END_DATE As Variant, _
Optional ByVal PERIOD_STR As String = "d", _
Optional ByVal ELEMENT_STR As String = "DO", _
Optional ByVal HEADERS_FLAG As Boolean = False, _
Optional ByVal ADJUST_FLAG As Boolean = False, _
Optional ByVal RESORT_FLAG As Boolean = True)

Dim i As Integer
Dim j As Integer
Dim k As Integer

Dim NROWS As Integer

Dim END_DAY_INT As Integer
Dim END_YEAR_INT As Integer
Dim END_MONTH_INT As Integer

Dim START_DAY_INT As Integer
Dim START_YEAR_INT As Integer
Dim START_MONTH_INT As Integer

Dim DATA_GROUP() As YAHOO_DATA_OBJ
Dim TEMP_VECTOR As Variant
Dim TEMP_MATRIX As Variant
Dim TICKERS_VECTOR As Variant

On Error GoTo ERROR_LABEL

START_YEAR_INT = Year(START_DATE)
START_MONTH_INT = Month(START_DATE)
START_DAY_INT = Day(START_DATE)

END_YEAR_INT = Year(END_DATE)
END_MONTH_INT = Month(END_DATE)
END_DAY_INT = Day(END_DATE)

If IsArray(TICKERS_RNG) = True Then
    TICKERS_VECTOR = TICKERS_RNG
    If UBound(TICKERS_VECTOR, 1) = 1 Then
        TICKERS_VECTOR = MATRIX_TRANSPOSE_FUNC(TICKERS_VECTOR)
    End If
Else
    TEMP_MATRIX = YAHOO_HISTORICAL_DATA_MATRIX_FUNC(TICKERS_RNG, START_YEAR_INT, START_MONTH_INT, START_DAY_INT, END_YEAR_INT, END_MONTH_INT, END_DAY_INT, PERIOD_STR, ELEMENT_STR, HEADERS_FLAG, ADJUST_FLAG, RESORT_FLAG, 0, 0)
    GoTo 1984
End If
NROWS = UBound(TICKERS_VECTOR)

j = 0: k = 1
For i = 1 To NROWS
    TEMP_MATRIX = YAHOO_HISTORICAL_DATA_MATRIX_FUNC(TICKERS_VECTOR(i, 1), START_YEAR_INT, START_MONTH_INT, START_DAY_INT, END_YEAR_INT, END_MONTH_INT, END_DAY_INT, PERIOD_STR, ELEMENT_STR, False, ADJUST_FLAG, False, 0, 0)
    If IS_2D_ARRAY_FUNC(TEMP_MATRIX) = False Then: GoTo 1983
    If UBound(TEMP_MATRIX, 1) >= k Then: k = UBound(TEMP_MATRIX, 1)
    j = j + 1
    ReDim Preserve DATA_GROUP(1 To j)
    DATA_GROUP(j).iSymbol = TICKERS_VECTOR(i, 1)
    DATA_GROUP(j).iData = TEMP_MATRIX
1983:
Next i

TEMP_MATRIX = YAHOO_HISTORICAL_DATA_MATCH_DATES_FUNC(DATA_GROUP(), k, RESORT_FLAG)

'--------------------------------------------------------------------------------------
If HEADERS_FLAG = True Then
'--------------------------------------------------------------------------------------
    TEMP_VECTOR = TEMP_MATRIX
    ReDim TEMP_MATRIX(0 To UBound(TEMP_VECTOR, 1), 1 To UBound(TEMP_VECTOR, 2))
    For j = 1 To UBound(TEMP_VECTOR, 2)
        If j > 1 Then TEMP_MATRIX(0, j) = DATA_GROUP(j - 1).iSymbol Else: TEMP_MATRIX(0, j) = "DATE"
        For i = 1 To UBound(TEMP_VECTOR, 1)
            TEMP_MATRIX(i, j) = TEMP_VECTOR(i, j)
        Next i
    Next j
'--------------------------------------------------------------------------------------
End If
'--------------------------------------------------------------------------------------

1984:
YAHOO_HISTORICAL_DATA_SERIES2_FUNC = TEMP_MATRIX
    
Exit Function
ERROR_LABEL:
YAHOO_HISTORICAL_DATA_SERIES2_FUNC = Err.number
End Function



Function YAHOO_HISTORICAL_DATA_SERIE_DATES_FUNC(ByVal TICKER_STR As String, _
ParamArray DATES_RNG() As Variant)

Dim h As Integer
Dim i As Long
Dim j As Long
Dim k As Long

Dim ii As Long
Dim jj As Long

Dim SDATE_VAL As Variant
Dim EDATE_VAL As Variant
Dim CDATE_VAL As Variant

Dim DCELL As Excel.Range

Dim DATA_ARR As Variant
Dim DATES_ARR As Variant
Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

'-----------------------------------------------------------------------------------------------------------*
' Function to return prices for multiple historical dates
'YAHOO_HISTORICAL_DATA_SERIE_DATES_FUNC("MMM",DATE(2007,1,1),DATE(2007,3,4))
'YAHOO_HISTORICAL_DATA_SERIE_DATES_FUNC("MMM","1/1/2007")
'YAHOO_HISTORICAL_DATA_SERIE_DATES_FUNC("MMM",C4:D4)
'YAHOO_HISTORICAL_DATA_SERIE_DATES_FUNC("MMM",DATE(2007,1,1),DATE(2007,3,4),C4:D4)
'-----------------------------------------------------------------------------------------------------------*
'Extract passed dates from parameters and/or ranges
'-----------------------------------------------------------------------------------------------------------*

GoSub DATES_LINE
'Get historical data and extract requested data
DATA_MATRIX = YAHOO_HISTORICAL_DATA_MATRIX_FUNC(TICKER_STR, 0, 0, 0, 0, 0, 0, "d", "DA", False, False, False, h, 2)

ReDim TEMP_MATRIX(1 To k)
For i = 1 To k
    If DATES_ARR(i) = "" Or DATES_ARR(i) > CDATE_VAL Then
        TEMP_MATRIX(i) = CVErr(xlErrNA)
    ElseIf DATES_ARR(i) > DATA_MATRIX(1, 1) Then
        If DATES_ARR(i) = CDATE_VAL Then
            DATA_ARR = MATRIX_YAHOO_QUOTES_FUNC(TICKER_STR, "l1", "", "", False)
            TEMP_MATRIX(i) = DATA_ARR(1, 1)
        Else
            TEMP_MATRIX(i) = DATA_MATRIX(1, 2)
        End If
    Else
        ii = 1: jj = h
        Do
            j = Int((jj + ii) / 2)
            If DATES_ARR(i) = DATA_MATRIX(j, 1) Then
                TEMP_MATRIX(i) = DATA_MATRIX(j, 2)
                Exit Do
            ElseIf ii = jj - 1 Then
                If DATA_MATRIX(jj, 2) <> "" Then
                    TEMP_MATRIX(i) = DATA_MATRIX(jj, 2)
                Else
                    TEMP_MATRIX(i) = CVErr(xlErrNA)
                End If
                Exit Do
            Else
                If DATES_ARR(i) > DATA_MATRIX(j, 1) Or DATA_MATRIX(j, 1) = "" Then
                    jj = j
                Else
                    ii = j
                End If
            End If
        Loop While True
    End If
Next i

'Return data
YAHOO_HISTORICAL_DATA_SERIE_DATES_FUNC = TEMP_MATRIX

Exit Function
'----------------------------------------------------------------------------------------
DATES_LINE:
'----------------------------------------------------------------------------------------
    CDATE_VAL = Now
    CDATE_VAL = DateSerial(Year(CDATE_VAL), Month(CDATE_VAL), Day(CDATE_VAL))
    EDATE_VAL = CDATE_VAL
    k = 0: ReDim DATES_ARR(1 To 1)
    For i = LBound(DATES_RNG) To UBound(DATES_RNG)
        Select Case VarType(DATES_RNG(i))
        Case vbDate, vbDouble
            SDATE_VAL = DATES_RNG(i)
            GoSub ARRAY_LINE
        Case vbString
            If IsDate(DATES_RNG(i)) Then
                SDATE_VAL = DateValue(DATES_RNG(i))
                GoSub ARRAY_LINE
            Else
                SDATE_VAL = ""
                GoSub ARRAY_LINE
            End If
        Case Is >= 8192
            For Each DCELL In DATES_RNG(i)
                Select Case VarType(DCELL.value)
                Case vbDate
                    SDATE_VAL = DCELL.value
                    GoSub ARRAY_LINE
                Case Else
                    SDATE_VAL = ""
                    GoSub ARRAY_LINE
                End Select
            Next DCELL
        Case Else
            SDATE_VAL = ""
            GoSub ARRAY_LINE
        End Select
    Next i
    h = Int(CDATE_VAL - EDATE_VAL + 3)
'----------------------------------------------------------------------------------------
Return
'----------------------------------------------------------------------------------------
ARRAY_LINE:
'----------------------------------------------------------------------------------------
    k = k + 1
    ReDim Preserve DATES_ARR(1 To k)
    If SDATE_VAL = "" Then
        DATES_ARR(k) = ""
    ElseIf Year(SDATE_VAL) < 1928 Or Year(SDATE_VAL) > Year(CDATE_VAL) Then
        DATES_ARR(k) = ""
    Else
        DATES_ARR(k) = SDATE_VAL
        If SDATE_VAL < EDATE_VAL Then EDATE_VAL = SDATE_VAL
    End If
'----------------------------------------------------------------------------------------
Return
'----------------------------------------------------------------------------------------
ERROR_LABEL:
YAHOO_HISTORICAL_DATA_SERIE_DATES_FUNC = Err.number
End Function

'// PERFECT

Function YAHOO_HISTORICAL_DATA_SERIES_DATES_FUNC(ByRef TICKERS_RNG As Variant, _
ByRef DATES_RNG As Variant, _
Optional ByVal ELEMENT_INT As Integer = 6, _
Optional ByVal HEADERS_FLAG As Boolean = True)

'ELEMENT_INT (1)Open; (2)High; (3)Low; (4)Close; (5) Volume; (6)Adj.Close

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long

Dim ii As Long
Dim jj As Long

Dim SROW As Long
Dim NROWS As Long
Dim NCOLUMNS As Long

Dim DATE_VAL As Variant
Dim CDATE_VAL As Variant 'Current Date
Dim SDATE_VAL As Variant 'Starting Date
Dim EDATE_VAL As Variant 'Ending Date

Dim ELEMENT_STR As String
Dim CURRENT_FLAG As Boolean
Dim RESET_FLAG As Boolean

Dim DATES_VECTOR As Variant
Dim TICKERS_VECTOR As Variant

Dim DATA_VECTOR As Variant
Dim DATA_MATRIX As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

If IsArray(DATES_RNG) = True Then
    DATES_VECTOR = DATES_RNG
    If UBound(DATES_VECTOR, 1) = 1 Then
        DATES_VECTOR = MATRIX_TRANSPOSE_FUNC(DATES_VECTOR)
    End If
Else
    ReDim DATES_VECTOR(1 To 1, 1 To 1)
    DATES_VECTOR(1, 1) = DATES_RNG
End If
DATES_VECTOR = MATRIX_QUICK_SORT_FUNC(DATES_VECTOR, 1, 1)
ii = UBound(DATES_VECTOR, 1)

If IsArray(TICKERS_RNG) = True Then
    TICKERS_VECTOR = TICKERS_RNG
    If UBound(TICKERS_VECTOR, 2) = 1 Then
        TICKERS_VECTOR = MATRIX_TRANSPOSE_FUNC(TICKERS_VECTOR)
    End If
Else
    ReDim TICKERS_VECTOR(1 To 1, 1 To 1)
    TICKERS_VECTOR(1, 1) = TICKERS_RNG
End If
jj = UBound(TICKERS_VECTOR, 2)

CDATE_VAL = Now
CDATE_VAL = DateSerial(Year(CDATE_VAL), Month(CDATE_VAL), Day(CDATE_VAL))
SDATE_VAL = DateSerial(2050, 1, 1)
EDATE_VAL = DateSerial(1950, 1, 1)
CURRENT_FLAG = False
For i = 1 To ii
    If IsDate(DATES_VECTOR(i, 1)) Then
        DATE_VAL = DATES_VECTOR(i, 1)
        DATE_VAL = DateSerial(Year(DATE_VAL), Month(DATE_VAL), Day(DATE_VAL))
        If DATE_VAL = CDATE_VAL Then: CURRENT_FLAG = True
        If DATE_VAL < SDATE_VAL Then: SDATE_VAL = DATE_VAL
        If DATE_VAL > EDATE_VAL Then: EDATE_VAL = DATE_VAL
    End If
1983:
Next i
If SDATE_VAL = DateSerial(2050, 1, 1) Or EDATE_VAL = DateSerial(1950, 1, 1) Then: GoTo ERROR_LABEL
If CURRENT_FLAG = True Then
    Select Case ELEMENT_INT
    Case 1
        ELEMENT_STR = "open"
    Case 2
        ELEMENT_STR = "high"
    Case 3
        ELEMENT_STR = "low"
    Case 4
        ELEMENT_STR = "last trade"
    Case 5
        ELEMENT_STR = "volume"
    Case Else
        ELEMENT_STR = "last trade"
    End Select
    DATA_VECTOR = YAHOO_QUOTES_FUNC(TICKERS_VECTOR, ELEMENT_STR, 0, False, "UNITED STATES")
End If
DATA_MATRIX = YAHOO_HISTORICAL_DATA_SERIES1_FUNC(TICKERS_RNG, SDATE_VAL, EDATE_VAL, ELEMENT_INT, "d", HEADERS_FLAG, True) 'Ascending
SROW = LBound(DATA_MATRIX, 1)
NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)
If HEADERS_FLAG = True Then
    ReDim TEMP_MATRIX(0 To ii, 1 To jj + 1)
    For j = 1 To NCOLUMNS: TEMP_MATRIX(0, j) = DATA_MATRIX(SROW, j): Next j
    h = 1
Else
    ReDim TEMP_MATRIX(1 To ii, 1 To jj + 1)
    h = 0
End If

For k = 1 To ii
    RESET_FLAG = True
    If CURRENT_FLAG = True And DATES_VECTOR(k, 1) = CDATE_VAL Then
        TEMP_MATRIX(k, 1) = CDATE_VAL
        For j = 2 To NCOLUMNS
            TEMP_MATRIX(k, j) = DATA_VECTOR(j - 1, 1)
            If TEMP_MATRIX(k, j) = 0 Then: TEMP_MATRIX(k, j) = CVErr(xlErrNA)
        Next j
        RESET_FLAG = False
    Else
        TEMP_MATRIX(k, 1) = DATES_VECTOR(k, 1)
        For i = SROW + h To NROWS
            If DATES_VECTOR(k, 1) = DATA_MATRIX(i, 1) Then
                For j = 2 To NCOLUMNS
                    TEMP_MATRIX(k, j) = DATA_MATRIX(i, j)
                    If TEMP_MATRIX(k, j) = 0 Then: TEMP_MATRIX(k, j) = CVErr(xlErrNA)
                Next j
                RESET_FLAG = False
                SROW = i + 1
                Exit For
            End If
        Next i
    End If
    If RESET_FLAG = True Then: For j = 2 To NCOLUMNS: TEMP_MATRIX(k, j) = CVErr(xlErrNA): Next j
Next k

YAHOO_HISTORICAL_DATA_SERIES_DATES_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
YAHOO_HISTORICAL_DATA_SERIES_DATES_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************



Function YAHOO_HISTORICAL_DATA_SERIES_PRICES_FUNC(ByVal TICKERS_RNG As Variant, _
ByVal SDATE_VAL As Date, _
ByVal EDATE_VAL As Date)
                                                
Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long
Dim m As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Const ELEMENT_STR As String = "01020304050607080910"
Dim TICKER_STR As String
Dim TICKERS_VECTOR As Variant
Dim DATA_ARR As Variant
Dim DATA_MATRIX As Variant

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
NCOLUMNS = CInt(Right(ELEMENT_STR, 2))
m = Len(ELEMENT_STR)

ReDim DATA_ARR(1 To NCOLUMNS)
ReDim TEMP_MATRIX(0 To NROWS, 1 To NCOLUMNS + 1)
TEMP_MATRIX(0, 1) = "TICKER"
If NCOLUMNS >= 1 Then: TEMP_MATRIX(0, 2) = "OPEN DATE"
If NCOLUMNS >= 2 Then: TEMP_MATRIX(0, 3) = "OPEN PRICE"
If NCOLUMNS >= 3 Then: TEMP_MATRIX(0, 4) = "HIGH DATE"
If NCOLUMNS >= 4 Then: TEMP_MATRIX(0, 5) = "HIGH PRICE"
If NCOLUMNS >= 5 Then: TEMP_MATRIX(0, 6) = "LOW DATE"
If NCOLUMNS >= 6 Then: TEMP_MATRIX(0, 7) = "LOW PRICE"
If NCOLUMNS >= 7 Then: TEMP_MATRIX(0, 8) = "CLOSE DATE"
If NCOLUMNS >= 8 Then: TEMP_MATRIX(0, 9) = "CLOSE PRICE"
If NCOLUMNS >= 9 Then: TEMP_MATRIX(0, 10) = "VOLUME"
If NCOLUMNS >= 10 Then: TEMP_MATRIX(0, 11) = "PREVIOUS CLOSE"

For k = 1 To NROWS
    TICKER_STR = TICKERS_VECTOR(k, 1)
    TEMP_MATRIX(k, 1) = TICKER_STR
    'h = 1: DATA_MATRIX = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, SDATE_VAL, EDATE_VAL, "DAILY", "DOHLCV", False, True, False)
    h = 0: DATA_MATRIX = YAHOO_HISTORICAL_DATA_MATRIX_FUNC(TICKER_STR, 0, 0, 0, 0, 0, 0, "d", "DOHLCV", False, True, False, CInt(Now - SDATE_VAL + 3), 6)
    If IsArray(DATA_MATRIX) = False Then: GoTo 1983
    l = UBound(DATA_MATRIX, 1)
    DATA_ARR(1) = ""     ' Date of open price
    DATA_ARR(2) = 0      ' Value of open price
    DATA_ARR(3) = ""     ' Date of high price
    DATA_ARR(4) = 0      ' Value of high price
    DATA_ARR(5) = ""     ' Date of low price
    DATA_ARR(6) = 0      ' Value of low price
    DATA_ARR(7) = ""     ' Date of closing price
    DATA_ARR(8) = 0      ' Value of closing price
    DATA_ARR(9) = 0      ' Total volume during period
    DATA_ARR(10) = 0     ' Value of previous closing price
    For i = 1 To l - h
        Select Case DATA_MATRIX(i, 1)
        Case Is > EDATE_VAL
        Case Is < SDATE_VAL: Exit For
        Case Else
            If DATA_ARR(8) = 0 Then
               DATA_ARR(3) = DATA_MATRIX(i, 1)  ' Latest date
               DATA_ARR(4) = DATA_MATRIX(i, 3)  ' Latest high
               DATA_ARR(5) = DATA_MATRIX(i, 1)  ' Latest date
               DATA_ARR(6) = DATA_MATRIX(i, 4)  ' Latest low
               DATA_ARR(7) = DATA_MATRIX(i, 1)  ' Latest date
               DATA_ARR(8) = DATA_MATRIX(i, 5)  ' Latest close
            End If
            DATA_ARR(1) = DATA_MATRIX(i, 1)     ' Earliest date
            DATA_ARR(2) = DATA_MATRIX(i, 2)     ' Earliest open
            DATA_ARR(10) = DATA_MATRIX(i + 1, 5)  ' Previous closing price
            DATA_ARR(9) = DATA_ARR(9) + DATA_MATRIX(i, 6)
            If DATA_ARR(6) > DATA_MATRIX(i, 4) Then
               DATA_ARR(5) = DATA_MATRIX(i, 1)  ' Date of lowest
               DATA_ARR(6) = DATA_MATRIX(i, 4)  ' Lower low
            End If
            If DATA_ARR(4) < DATA_MATRIX(i, 3) Then
               DATA_ARR(3) = DATA_MATRIX(i, 1)  ' Date of highest
               DATA_ARR(4) = DATA_MATRIX(i, 3)  ' Higher high
            End If
        End Select
    Next i
    For j = 1 To NCOLUMNS
        If 2 * j > m Then
           TEMP_MATRIX(k, j + 1) = ""
        Else
           i = CInt(Mid(ELEMENT_STR, 2 * j - 1, 2))
           TEMP_MATRIX(k, j + 1) = DATA_ARR(i)
        End If
    Next j
1983:
Next k
'Erase DATA_MATRIX: Erase DATA_ARR
YAHOO_HISTORICAL_DATA_SERIES_PRICES_FUNC = TEMP_MATRIX
   
Exit Function
ERROR_LABEL:
YAHOO_HISTORICAL_DATA_SERIES_PRICES_FUNC = Err.number
End Function
        
Function YAHOO_HISTORICAL_DATA_HIGH_DATE_FUNC(ByVal TICKER_STR As String, _
ByVal NO_DAYS As Long)

Dim i As Long
Dim NROWS As Long

Dim EDATE_VAL As Date
Dim SDATE_VAL As Date
Dim DATA_VAL As Double
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = YAHOO_HISTORICAL_DATA_MATRIX_FUNC(TICKER_STR, 0, 0, 0, 0, 0, 0, "d", "DH", False, True, False, CInt(NO_DAYS), 2)
NROWS = UBound(DATA_MATRIX, 1)
i = 1
SDATE_VAL = DATA_MATRIX(i, 1)
EDATE_VAL = SDATE_VAL: DATA_VAL = DATA_MATRIX(i, 2)
For i = 2 To NROWS
    If DATA_MATRIX(i, 1) < SDATE_VAL - NO_DAYS Then Exit For
    If DATA_MATRIX(i, 2) > DATA_VAL Then
       DATA_VAL = DATA_MATRIX(i, 2)
       EDATE_VAL = DATA_MATRIX(i, 1)
    End If
Next i

YAHOO_HISTORICAL_DATA_HIGH_DATE_FUNC = Array(EDATE_VAL, DATA_VAL)

Exit Function
ERROR_LABEL:
YAHOO_HISTORICAL_DATA_HIGH_DATE_FUNC = Err.number
End Function

Function YAHOO_HISTORICAL_DATA_LAST_PRICE_FUNC(ByVal TICKER_STR As String, _
ByVal DATE_VAL As Date)

Dim i As Long
Dim j As Long
Dim NROWS As Long
Dim DATA_VAL As Double
Dim DATA_MATRIX As Variant
'YAHOO_HISTORICAL_DATA_LAST_PRICE_FUNC("MMM",DATE(2007,1,1))
   
On Error GoTo ERROR_LABEL
   
j = 9999
DATA_MATRIX = YAHOO_HISTORICAL_DATA_MATRIX_FUNC(TICKER_STR, 0, 0, 0, 0, 0, 0, "d", "DA", False, True, False, CInt(j), 2)
NROWS = UBound(DATA_MATRIX, 1)
DATA_VAL = 0
For i = 1 To NROWS
    If DATA_MATRIX(i, 1) <= DATE_VAL Then
       DATA_VAL = DATA_MATRIX(i, 2)
       Exit For
    End If
Next i

YAHOO_HISTORICAL_DATA_LAST_PRICE_FUNC = DATA_VAL

Exit Function
ERROR_LABEL:
YAHOO_HISTORICAL_DATA_LAST_PRICE_FUNC = Err.number
End Function

Function YAHOO_HISTORICAL_DATA_HIGH_BETWEEN_FUNC(ByVal TICKER_STR As String, _
ByVal SDATE_VAL As Date, _
ByVal EDATE_VAL As Date)

'TICKER_STR: Ticker symbol (e.g "MMM"). You can also use a special value
'of "Header" to return a row of column headings for the returned data items.

'Start Date: Date to use as the starting date for the historical quotes
'date range; must be an EXCEL serial date value (e.g. NOW() or DATE(2006,10,15)).

'End Date: Date to use as the ending date for the historical quotes date
'range; must be an EXCEL serial date value (e.g. TODAY() or DATE(2007,10,15)).
   
Dim i As Long
Dim j As Long
Dim NROWS As Long

Dim DATA_VECTOR As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

' Checks for highest price between two dates
j = 9999
DATA_MATRIX = YAHOO_HISTORICAL_DATA_MATRIX_FUNC(TICKER_STR, 0, 0, 0, 0, 0, 0, "d", "DHOC", False, True, False, CInt(j), 4)

NROWS = UBound(DATA_MATRIX, 1)

ReDim DATA_VECTOR(1 To 1, 1 To 4)
DATA_VECTOR(1, 1) = 0    ' Value of high price
DATA_VECTOR(1, 2) = ""   ' Day of high price
DATA_VECTOR(1, 3) = 0    ' Starting price
DATA_VECTOR(1, 4) = 0    ' Ending price

For i = 1 To NROWS
    Select Case DATA_MATRIX(i, 1)
    Case Is > EDATE_VAL
    Case Is < SDATE_VAL
        Exit For
    Case Else
         If DATA_MATRIX(i, 1) = SDATE_VAL Then DATA_VECTOR(1, 3) = DATA_MATRIX(i, 3)
         If DATA_MATRIX(i, 1) = EDATE_VAL Then DATA_VECTOR(1, 4) = DATA_MATRIX(i, 4)
         If DATA_MATRIX(i, 2) > DATA_VECTOR(1, 1) Then
            DATA_VECTOR(1, 1) = DATA_MATRIX(i, 2)
            DATA_VECTOR(1, 2) = DATA_MATRIX(i, 1)
         End If
    End Select
Next i

YAHOO_HISTORICAL_DATA_HIGH_BETWEEN_FUNC = DATA_VECTOR

'In most cases, this will need to be an array-entered formula. To
'array-enter a formula in EXCEL, first highlight the range of cells
'where you would like the returned data to appear -- the number of
'rows for the range should be the number of periods of data you are
'requesting from the function, while the number of columns for the
'range should be the number of data items you are requesting for each
'date from the function. Next, enter your formula and then press
'Ctrl-Shift-Enter.

Exit Function
ERROR_LABEL:
YAHOO_HISTORICAL_DATA_HIGH_BETWEEN_FUNC = Err.number
End Function


Function YAHOO_HISTORICAL_DATA_LOW_BETWEEN_FUNC(ByVal TICKER_STR As String, _
ByVal SDATE_VAL As Date, _
ByVal EDATE_VAL As Date)

Dim i As Long
Dim j As Long
Dim NROWS As Long

Dim DATA_VECTOR As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

j = 9999
DATA_MATRIX = YAHOO_HISTORICAL_DATA_MATRIX_FUNC(TICKER_STR, 0, 0, 0, 0, 0, 0, "d", "DLOC", False, True, False, CInt(j), 4)
NROWS = UBound(DATA_MATRIX, 1)
' Checks for highest price between two dates
ReDim DATA_VECTOR(1 To 1, 1 To 4)
DATA_VECTOR(1, 1) = 2 ^ 52 ' Value of low price
DATA_VECTOR(1, 2) = ""     ' Day of low price
DATA_VECTOR(1, 3) = 0      ' Starting price
DATA_VECTOR(1, 4) = 0      ' Ending price
For i = 1 To NROWS
    Select Case DATA_MATRIX(i, 1)
       Case Is > EDATE_VAL
       Case Is < SDATE_VAL: Exit For
       Case Else
            If DATA_VECTOR(1, 4) = 0 Then DATA_VECTOR(1, 4) = DATA_MATRIX(i, 4)
            DATA_VECTOR(1, 3) = DATA_MATRIX(i, 3)
            If DATA_MATRIX(i, 2) < DATA_VECTOR(1, 1) Then
               DATA_VECTOR(1, 1) = DATA_MATRIX(i, 2)
               DATA_VECTOR(1, 2) = DATA_MATRIX(i, 1)
            End If
    End Select
Next i
YAHOO_HISTORICAL_DATA_LOW_BETWEEN_FUNC = DATA_VECTOR

Exit Function
ERROR_LABEL:
YAHOO_HISTORICAL_DATA_LOW_BETWEEN_FUNC = Err.number
End Function


Function YAHOO_HISTORICAL_DATA_MATRIX_FUNC(ByVal TICKER_STR As String, _
Optional ByVal START_YEAR_INT As Integer = 0, _
Optional ByVal START_MONTH_INT As Integer = 0, _
Optional ByVal START_DAY_INT As Integer = 0, _
Optional ByVal END_YEAR_INT As Integer = 0, _
Optional ByVal END_MONTH_INT As Integer = 0, _
Optional ByVal END_DAY_INT As Integer = 0, _
Optional ByRef PERIOD_STR As String = "d", _
Optional ByRef ELEMENT_STR As String = "DOHLCVA", _
Optional ByVal HEADERS_FLAG As Boolean = True, _
Optional ByVal ADJUST_FLAG As Boolean = False, _
Optional ByVal RESORT_FLAG As Boolean = False, _
Optional ByRef NROWS As Integer = 0, _
Optional ByRef NCOLUMNS As Integer = 0)

Dim h As Integer
Dim l As Integer
Dim m As Integer

Dim hh As Integer
Dim ii As Integer
Dim jj As Integer
Dim kk As Integer
Dim ll As Integer
Dim mm As Integer
Dim nn As Integer
Dim oo As Integer
Dim pp As Integer

Dim SROW As Integer
Dim NSIZE As Integer

Dim DATA_STR As String
Dim TEMP_STR As String
Dim DATA_ARR As Variant
Dim TEMP_ARR As Variant
Dim TEMP_VAL As Variant

Dim ERROR_STR As String
Dim SRC_URL_STR As String
    
On Error GoTo ERROR_LABEL

If TICKER_STR = "None" Or TICKER_STR = "" Then 'Null Return Item
    ReDim DATA_ARR(1 To 1, 1 To 1)
    DATA_ARR(1, 1) = "None"
    YAHOO_HISTORICAL_DATA_MATRIX_FUNC = DATA_ARR
    Exit Function
End If

PERIOD_STR = LCase(PERIOD_STR)

h = 0
If ADJUST_FLAG = True Then: h = 1

SRC_URL_STR = YAHOO_HISTORICAL_DATA_URL_FUNC(TICKER_STR, START_YEAR_INT, _
START_MONTH_INT, START_DAY_INT, END_YEAR_INT, END_MONTH_INT, _
END_DAY_INT, PERIOD_STR, NROWS, ERROR_STR)

If ERROR_STR <> "" Then: GoTo ERROR_LABEL
DATA_STR = SAVE_WEB_DATA_PAGE_FUNC(SRC_URL_STR, 0, False, 0, False)

DATA_ARR = Split(DATA_STR, Chr(10), -1, vbBinaryCompare) 'Parse web quotes
SROW = LBound(DATA_ARR)
If NROWS = 0 Then: NROWS = UBound(DATA_ARR) - 1

ELEMENT_STR = UCase(ELEMENT_STR) 'Determine items needed
If PERIOD_STR = "v" Then
   If InStr(ELEMENT_STR, "T") > 0 Then: hh = 1
   ii = 1 + hh
   jj = 2 + hh
   kk = 0
   ll = 0
   mm = 0
   nn = 0
   oo = 0
   pp = 0
Else
   jj = 0 'iDiv
   hh = InStr(ELEMENT_STR, "T") 'iTick
   ii = InStr(ELEMENT_STR, "D") 'iDate
   kk = InStr(ELEMENT_STR, "O") 'iOpen
   ll = InStr(ELEMENT_STR, "H") 'iHigh
   mm = InStr(ELEMENT_STR, "L") 'iLow
   nn = InStr(ELEMENT_STR, "C") 'iClos
   oo = InStr(ELEMENT_STR, "V") 'iVol
   pp = InStr(ELEMENT_STR, "A") 'iAdjC
End If

If NCOLUMNS <> 0 Then
    If hh > NCOLUMNS Then hh = 0
    If ii > NCOLUMNS Then ii = 0
    If jj > NCOLUMNS Then jj = 0
    If kk > NCOLUMNS Then kk = 0
    If ll > NCOLUMNS Then ll = 0
    If mm > NCOLUMNS Then mm = 0
    If nn > NCOLUMNS Then nn = 0
    If oo > NCOLUMNS Then oo = 0
    If pp > NCOLUMNS Then pp = 0
Else
    'NCOLUMNS = UBound(Split(DATA_ARR(SROW), ",")) + 1
    If hh > 0 Then NCOLUMNS = 1
    If ii > 0 Then NCOLUMNS = NCOLUMNS + 1
    If jj > 0 Then NCOLUMNS = NCOLUMNS + 1
    If kk > 0 Then NCOLUMNS = NCOLUMNS + 1
    If ll > 0 Then NCOLUMNS = NCOLUMNS + 1
    If mm > 0 Then NCOLUMNS = NCOLUMNS + 1
    If nn > 0 Then NCOLUMNS = NCOLUMNS + 1
    If oo > 0 Then NCOLUMNS = NCOLUMNS + 1
    If pp > 0 Then NCOLUMNS = NCOLUMNS + 1
End If
    

If HEADERS_FLAG = True Then
    ReDim DATA_MATRIX(0 To NROWS, 1 To NCOLUMNS)
    TEMP_ARR = Split(DATA_ARR(SROW), ",", -1, vbBinaryCompare)
    TEMP_STR = IIf(h = 1, "Adj. ", "")
    If hh > 0 Then DATA_MATRIX(0, hh) = "Ticker"
    If ii > 0 Then DATA_MATRIX(0, ii) = TEMP_ARR(SROW)
    If jj > 0 Then DATA_MATRIX(0, jj) = TEMP_ARR(SROW + 1)
    If kk > 0 Then DATA_MATRIX(0, kk) = TEMP_STR & TEMP_ARR(SROW + 1)
    If ll > 0 Then DATA_MATRIX(0, ll) = TEMP_STR & TEMP_ARR(SROW + 2)
    If mm > 0 Then DATA_MATRIX(0, mm) = TEMP_STR & TEMP_ARR(SROW + 3)
    If nn > 0 Then DATA_MATRIX(0, nn) = TEMP_STR & TEMP_ARR(SROW + 4)
    If oo > 0 Then DATA_MATRIX(0, oo) = TEMP_ARR(SROW + 5)
    If pp > 0 Then DATA_MATRIX(0, pp) = TEMP_ARR(SROW + 6)
Else
    ReDim DATA_MATRIX(1 To NROWS, 1 To NCOLUMNS)
End If


For l = 1 To NROWS: For m = 1 To NCOLUMNS: DATA_MATRIX(l, m) = "": Next m: Next l
NSIZE = IIf(NROWS < UBound(DATA_ARR) - 1, NROWS, UBound(DATA_ARR) - 1)

For l = 1 To NSIZE
    TEMP_ARR = Split(DATA_ARR(SROW + l), ",", -1, vbBinaryCompare)
    
    If hh > 0 Then: DATA_MATRIX(SROW + l, hh) = TICKER_STR
    If ii > 0 Then
         DATA_MATRIX(SROW + l, ii) = NORMALIZE_DATES_VECTOR_FUNC(CDate(TEMP_ARR(SROW)), PERIOD_STR)
    End If
    If jj > 0 Then DATA_MATRIX(SROW + l, jj) = CONVERT_STRING_NUMBER_FUNC(TEMP_ARR(SROW + 1))
    
    If PERIOD_STR <> "v" Then
       If h = 1 Then
         If Val(TEMP_ARR(SROW + 4)) <> 0 Then
             TEMP_VAL = CONVERT_STRING_NUMBER_FUNC(TEMP_ARR(SROW + 6)) / CONVERT_STRING_NUMBER_FUNC(TEMP_ARR(SROW + 4))
         Else
             TEMP_VAL = 1 '0
         End If
       Else
         TEMP_VAL = 1
       End If
       If kk > 0 Then
         DATA_MATRIX(l, kk) = CONVERT_STRING_NUMBER_FUNC(TEMP_ARR(SROW + 1)) * TEMP_VAL
       End If
       If ll > 0 Then
         DATA_MATRIX(l, ll) = CONVERT_STRING_NUMBER_FUNC(TEMP_ARR(SROW + 2)) * TEMP_VAL
       End If
       If mm > 0 Then
         DATA_MATRIX(l, mm) = CONVERT_STRING_NUMBER_FUNC(TEMP_ARR(SROW + 3)) * TEMP_VAL
       End If
       If nn > 0 Then
         DATA_MATRIX(l, nn) = CONVERT_STRING_NUMBER_FUNC(TEMP_ARR(SROW + 4)) * TEMP_VAL
       End If
       If oo > 0 Then
         DATA_MATRIX(l, oo) = CONVERT_STRING_NUMBER_FUNC(TEMP_ARR(SROW + 5))
       End If
       If pp > 0 Then
         DATA_MATRIX(l, pp) = CONVERT_STRING_NUMBER_FUNC(TEMP_ARR(SROW + 6))
       End If
    End If
Next l

If RESORT_FLAG = True And PERIOD_STR <> "v" Then
    Call YAHOO_HISTORICAL_DATA_SORT_DATA_FUNC(DATA_MATRIX, SROW, NSIZE)
End If
    
YAHOO_HISTORICAL_DATA_MATRIX_FUNC = DATA_MATRIX

   
Exit Function
ERROR_LABEL:
YAHOO_HISTORICAL_DATA_MATRIX_FUNC = ERROR_STR
End Function


Function YAHOO_HISTORICAL_DATA_MATCH_DATES_FUNC( _
ByRef DATA_GROUP() As YAHOO_DATA_OBJ, _
ByRef NROWS As Integer, _
Optional ByVal RESORT_FLAG As Boolean = True)

Dim i As Integer
Dim j As Integer
Dim k As Integer

Dim NCOLUMNS As Integer
Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL
        
NCOLUMNS = UBound(DATA_GROUP()) * 2
ReDim DATA_MATRIX(1 To NROWS, 1 To NCOLUMNS)
k = 1
For j = 1 To NCOLUMNS Step 2
    TEMP_MATRIX = DATA_GROUP(k).iData
    NROWS = UBound(TEMP_MATRIX, 1)
    For i = 1 To NROWS
        DATA_MATRIX(i, j) = TEMP_MATRIX(i, 1)
        DATA_MATRIX(i, j + 1) = TEMP_MATRIX(i, 2)
    Next i
    k = k + 1
Next j
YAHOO_HISTORICAL_DATA_MATCH_DATES_FUNC = MATCH_DATES_VECTOR1_FUNC(DATA_MATRIX, IIf(RESORT_FLAG = True, 0, 1))

Exit Function
ERROR_LABEL:
YAHOO_HISTORICAL_DATA_MATCH_DATES_FUNC = Err.number
End Function

Function YAHOO_HISTORICAL_DATA_SORT_DATA_FUNC(ByRef DATA_RNG As Variant, _
ByVal SROW As Integer, _
ByVal NROWS As Integer)

Dim i As Integer
Dim j As Integer
Dim k As Integer

Dim SCOLUMN As Integer
Dim NCOLUMNS As Integer

Dim TEMP_VAL As Variant

On Error GoTo ERROR_LABEL

SCOLUMN = LBound(DATA_RNG, 2)
NCOLUMNS = UBound(DATA_RNG, 2)

'------------------------------------------------------------
i = 1 + SROW
j = NROWS + SROW
'------------------------------------------------------------
Do While i < j
'------------------------------------------------------------
    For k = SCOLUMN To NCOLUMNS
        TEMP_VAL = DATA_RNG(i, k)
        DATA_RNG(i, k) = DATA_RNG(j, k)
        DATA_RNG(j, k) = TEMP_VAL
    Next k
    i = i + 1
    j = j - 1
'------------------------------------------------------------
Loop
'------------------------------------------------------------

Exit Function
ERROR_LABEL:
YAHOO_HISTORICAL_DATA_SORT_DATA_FUNC = Err.number
End Function


Function YAHOO_HISTORICAL_DATA_URL_FUNC(ByVal TICKER_STR As String, _
ByVal START_YEAR_INT As Integer, _
ByVal START_MONTH_INT As Integer, _
ByVal START_DAY_INT As Integer, _
ByVal END_YEAR_INT As Integer, _
ByVal END_MONTH_INT As Integer, _
ByVal END_DAY_INT As Integer, _
ByVal PERIOD_STR As String, _
Optional ByVal NROWS As Integer = 0, _
Optional ByRef ERROR_STR As String = "")

Dim SRC_URL_STR As String
Dim END_YEAR_VAL As Integer

On Error GoTo ERROR_LABEL

ERROR_STR = ""

If YAHOO_HISTORICAL_DATA_VALIDATE_DATES_FUNC(START_YEAR_INT, START_MONTH_INT, _
    START_DAY_INT, END_YEAR_INT, END_MONTH_INT, END_DAY_INT, _
    ERROR_STR) = False Then: GoTo ERROR_LABEL

'----------------------------------------------------------------------------------
If START_YEAR_INT = 0 Then
'----------------------------------------------------------------------------------
    Select Case PERIOD_STR
       Case "d"
            END_YEAR_VAL = Year(Date) - CInt(NROWS / 250) - 1
       Case "w"
            END_YEAR_VAL = Year(Date) - CInt(NROWS / 50) - 1
       Case "m"
            END_YEAR_VAL = Year(Date) - CInt(NROWS / 12) - 1
       Case "v"
            END_YEAR_VAL = Year(Date) - CInt(NROWS / 4) - 1
       Case Else
           ERROR_STR = "Invalid Period Requested"
           GoTo ERROR_LABEL
    End Select
'----------------------------------------------------------------------------------
End If
'----------------------------------------------------------------------------------

SRC_URL_STR = "http://ichart.finance.yahoo.com/table.csv?s="
SRC_URL_STR = SRC_URL_STR & TICKER_STR & _
       IIf(START_MONTH_INT = 0, "&a=0", "&a=" & (START_MONTH_INT - 1)) & _
       IIf(START_DAY_INT = 0, "&b=1", "&b=" & START_DAY_INT) & _
       IIf(START_YEAR_INT = 0, "&c=" & END_YEAR_VAL, "&c=" & START_YEAR_INT) & _
       IIf(END_MONTH_INT = 0, "", "&d=" & (END_MONTH_INT - 1)) & _
       IIf(END_DAY_INT = 0, "", "&e=" & END_DAY_INT) & _
       IIf(END_YEAR_INT = 0, "", "&f=" & END_YEAR_INT) & _
       "&g=" & PERIOD_STR & _
       "&ignore=.csv"

YAHOO_HISTORICAL_DATA_URL_FUNC = SRC_URL_STR

Exit Function
ERROR_LABEL:
YAHOO_HISTORICAL_DATA_URL_FUNC = ERROR_STR
End Function

'// PERFECT



Function YAHOO_HISTORICAL_DATA_VALIDATE_DATES_FUNC( _
ByRef START_YEAR_INT As Integer, _
ByRef START_MONTH_INT As Integer, _
ByRef START_DAY_INT As Integer, _
ByRef END_YEAR_INT As Integer, _
ByRef END_MONTH_INT As Integer, _
ByRef END_DAY_INT As Integer, _
ByRef ERROR_STR As String)

On Error GoTo ERROR_LABEL

ERROR_STR = ""
YAHOO_HISTORICAL_DATA_VALIDATE_DATES_FUNC = True

'------------------> Edit parameters
If START_YEAR_INT = 0 And _
   START_MONTH_INT = 0 And _
   START_DAY_INT = 0 And _
   END_YEAR_INT = 0 And _
   END_MONTH_INT = 0 And _
   END_DAY_INT = 0 Then
Else
   If START_YEAR_INT < 1900 Or START_YEAR_INT > 2100 Or _
         START_MONTH_INT < 1 Or START_MONTH_INT > 12 Or _
         START_DAY_INT < 1 Or START_DAY_INT > 31 Or _
         END_YEAR_INT < 1900 Or END_YEAR_INT > 2100 Or _
         END_MONTH_INT < 1 Or END_MONTH_INT > 12 Or _
         END_DAY_INT < 1 Or END_DAY_INT > 31 Or _
             START_YEAR_INT & _
             Right("0" & START_MONTH_INT, 2) & _
             Right("0" & START_DAY_INT, 2) > _
             END_YEAR_INT & _
             Right("0" & END_MONTH_INT, 2) & _
             Right("0" & END_DAY_INT, 2) Then
             
             ERROR_STR = "Something wrong with dates -- asked for " & _
                    START_YEAR_INT & "/" & _
                    START_MONTH_INT & "/" & _
                    START_DAY_INT & " thru " & _
                    END_YEAR_INT & "/" & _
                    END_MONTH_INT & "/" & _
                    END_DAY_INT
             YAHOO_HISTORICAL_DATA_VALIDATE_DATES_FUNC = False
             Exit Function
    End If
End If


Exit Function
ERROR_LABEL:
YAHOO_HISTORICAL_DATA_VALIDATE_DATES_FUNC = False
End Function

Function YAHOO_HISTORICAL_DATA_PERIOD_STRING_FUNC(ByVal PERIOD_STR As String)

On Error GoTo ERROR_LABEL

Select Case LCase(PERIOD_STR)
Case "d", "daily": YAHOO_HISTORICAL_DATA_PERIOD_STRING_FUNC = "d"
Case "w", "weekly": YAHOO_HISTORICAL_DATA_PERIOD_STRING_FUNC = "w"
Case "m", "monthly": YAHOO_HISTORICAL_DATA_PERIOD_STRING_FUNC = "m"
Case Else: YAHOO_HISTORICAL_DATA_PERIOD_STRING_FUNC = "v"
End Select

Exit Function
ERROR_LABEL:
YAHOO_HISTORICAL_DATA_PERIOD_STRING_FUNC = Err.number
End Function


Function YAHOO_HISTORICAL_DATA_ELEMENT_STRING_FUNC( _
ByVal ELEMENT_STR As String)

On Error GoTo ERROR_LABEL

Select Case UCase(ELEMENT_STR)
Case "DATE": YAHOO_HISTORICAL_DATA_ELEMENT_STRING_FUNC = 1
Case "OPEN": YAHOO_HISTORICAL_DATA_ELEMENT_STRING_FUNC = "2,do"
Case "HIGH": YAHOO_HISTORICAL_DATA_ELEMENT_STRING_FUNC = "3,dh"
Case "LOW": YAHOO_HISTORICAL_DATA_ELEMENT_STRING_FUNC = "4,dl"
Case "CLOSE": YAHOO_HISTORICAL_DATA_ELEMENT_STRING_FUNC = "5,dc"
Case "VOLUME": YAHOO_HISTORICAL_DATA_ELEMENT_STRING_FUNC = "6,dv"
Case "ADJ. CLOSE": YAHOO_HISTORICAL_DATA_ELEMENT_STRING_FUNC = "7,da"
Case "DIVIDENDS": YAHOO_HISTORICAL_DATA_ELEMENT_STRING_FUNC = "2"
End Select

Exit Function
ERROR_LABEL:
YAHOO_HISTORICAL_DATA_ELEMENT_STRING_FUNC = Err.number
End Function
