Attribute VB_Name = "WEB_SERVICE_YAHOO_OPTION_LIBR"

'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'------------------------------------------------------------------------------------
Private PUB_YAHOO_OPTIONS_HASH_OBJ As clsTypeHash
'------------------------------------------------------------------------------------

Function YAHOO_OPTION_QUOTES_CURRENT_FUNC(ByRef TICKERS_RNG As Variant, _
ByRef STRIKE_RNG As Variant)
'EXAMPLE:
'TICKERS_RNG = IBM, IBM, IBM
'STRIKE_RNG = 110, 120, 130

Dim i As Integer
Dim j As Integer
Dim k As Variant
Dim l As Integer
Dim NROWS As Integer
Dim NCOLUMNS As Integer
Dim CHR_STR As String
Dim REF_STR As String
Dim SRC_URL_STR As String
Dim HEADINGS_STR As String

Dim TEMP_MATRIX As Variant
Dim STRIKE_VECTOR As Variant
Dim TICKERS_VECTOR As Variant

On Error GoTo ERROR_LABEL

If IsArray(STRIKE_RNG) = True Then
    STRIKE_VECTOR = STRIKE_RNG
    If UBound(STRIKE_VECTOR, 1) = 1 Then
        STRIKE_VECTOR = MATRIX_TRANSPOSE_FUNC(STRIKE_VECTOR)
    End If
Else
    ReDim STRIKE_VECTOR(1 To 1, 1 To 1)
    STRIKE_VECTOR(1, 1) = STRIKE_RNG
End If
NROWS = UBound(STRIKE_VECTOR, 1)

If IsArray(TICKERS_RNG) = True Then
    TICKERS_VECTOR = TICKERS_RNG
    If UBound(TICKERS_VECTOR, 1) = 1 Then
        TICKERS_VECTOR = MATRIX_TRANSPOSE_FUNC(TICKERS_VECTOR)
    End If
Else
    ReDim TICKERS_VECTOR(1 To NROWS, 1 To 1)
    For i = 1 To NROWS
        TICKERS_VECTOR(i, 1) = TICKERS_RNG
    Next i
End If

If UBound(TICKERS_VECTOR, 1) <> NROWS Then: GoTo ERROR_LABEL
HEADINGS_STR = "Ticker,Strike Price," & _
               "Current Month,Later Month,Current Month Symbol (Call),Later Month Symbol (Call),Current Month Last Price (Call),Later Month Last Price (Call),Current Month Bid Price (Call),Later Month Bid Price (Call),Current Month Ask Price (Call),Later Month Ask Price (Call),Current Month Volume (Call),Later Month Volume (Call),Current Month Open Interest (Call),Later Month Open Interest (Call)," & _
               "Current Month Symbol (Put),Later Month Symbol (Put),Current Month Last Price (Put),Later Month Last Price (Put),Current Month Bid Price (Put),Later Month Bid Price (Put),Current Month Ask Price (Put),Later Month Ask Price (Put),Current Month Volume (Put),Later Month Volume (Put),Current Mon Open Interest (Put),Later Month Open Interest (Put),"
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

REF_STR = ">Expires"
CHR_STR = "&k="
SRC_URL_STR = "http://finance.yahoo.com/q/op?s="

k = Array(1, 1, 1, 2, 2, 1, 2, 2, 3, 1, 3, 2, 5, 1, 5, 2, 6, 1, 6, 2, 7, 1, 7, 2, 8, 1, 8, 2, _
                      2, 1, 2, 2, 3, 1, 3, 2, 5, 1, 5, 2, 6, 1, 6, 2, 7, 1, 7, 2, 8, 1, 8, 2)
'----------------------------------------------------------------------------------
For i = 1 To NROWS
'----------------------------------------------------------------------------------
    TEMP_MATRIX(i, 1) = TICKERS_VECTOR(i, 1)
    TEMP_MATRIX(i, 2) = STRIKE_VECTOR(i, 1)
    l = 1
    For j = 3 To NCOLUMNS
        If j <= (NCOLUMNS / 2 + 2) Then
            TEMP_MATRIX(i, j) = RETRIEVE_WEB_DATA_CELL_FUNC(SRC_URL_STR & TICKERS_VECTOR(i, 1) & CHR_STR & STRIKE_VECTOR(i, 1), k(l + 0), REF_STR, , , , k(l + 1))
        Else
            TEMP_MATRIX(i, j) = RETRIEVE_WEB_DATA_CELL_FUNC(SRC_URL_STR & TICKERS_VECTOR(i, 1) & CHR_STR & STRIKE_VECTOR(i, 1), k(l + 0), REF_STR, REF_STR, , , k(l + 1))
        End If
        If TEMP_MATRIX(i, j) = "Error" Then: TEMP_MATRIX(i, j) = "N/A"
        l = l + 2
    Next j
'----------------------------------------------------------------------------------
Next i
'----------------------------------------------------------------------------------


YAHOO_OPTION_QUOTES_CURRENT_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
YAHOO_OPTION_QUOTES_CURRENT_FUNC = Err.number
End Function


Function YAHOO_OPTIONS_EXPIRY_STRIKE_FUNC(Optional ByVal SYMBOL_STR As String = "AAPL", _
Optional ByVal VERSION As Integer = 0)

Dim h() As Long

Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim CHR_STR As String
Dim URL_STR As String
Dim KEY_STR As String
Dim ITEM_STR As String
Dim LINE_STR As String
Dim DATA_STR As String

Dim COLLECT_OBJ As Collection
Dim TEMP_VECTOR As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

GoSub LOAD_LINE

Select Case VERSION
Case 0
    YAHOO_OPTIONS_EXPIRY_STRIKE_FUNC = TEMP_VECTOR
Case Else
    GoSub MATRIX_LINE
    YAHOO_OPTIONS_EXPIRY_STRIKE_FUNC = TEMP_MATRIX
End Select

Exit Function
'------------------------------------------------------------------------------------------------------------------------------------
LOAD_LINE:
'------------------------------------------------------------------------------------------------------------------------------------
    
    CHR_STR = "m=2"
    URL_STR = "http://finance.yahoo.com/q/os?s="
    DATA_STR = SAVE_WEB_DATA_PAGE_FUNC(URL_STR & SYMBOL_STR, 0, True, 0, True, CHR_STR, 32767, 0)
    i = 0: i = InStr(1, DATA_STR, CHR_STR)
    If i = 0 Then: GoTo ERROR_LABEL
    j = 1: ReDim h(1 To j)
    h(j) = i
    i = InStr(i + 1, DATA_STR, CHR_STR)
    Do While i <> 0
        j = j + 1: ReDim Preserve h(1 To j): h(j) = i
        i = InStr(i + 1, DATA_STR, CHR_STR)
    Loop
    If VERSION = 0 Then
        ReDim TEMP_VECTOR(1 To j, 1 To 3)
    Else
        Set COLLECT_OBJ = New Collection
    End If
    On Error Resume Next
    For i = 1 To j
        ITEM_STR = Replace(Mid(DATA_STR, h(i), 9), "m=", "")
        GoSub RETRIEVE_LINE
        If VERSION = 0 Then: GoSub VECTOR_LINE
    Next i
    On Error GoTo ERROR_LABEL
'------------------------------------------------------------------------------------------------------------------------------------
Return
'------------------------------------------------------------------------------------------------------------------------------------
RETRIEVE_LINE:
'------------------------------------------------------------------------------------------------------------------------------------
    k = 0
    Do
        k = k + 1
        KEY_STR = Trim(RETRIEVE_WEB_DATA_CELL_FUNC(URL_STR & SYMBOL_STR & "&m=" & ITEM_STR, 8, "Options Expiring", "Symbol", , , k, "</table", , ""))
        If VERSION > 0 Then: GoSub KEY_LINE
    Loop Until KEY_STR = ""
    k = k - 1
'------------------------------------------------------------------------------------------------------------------------------------
Return
'------------------------------------------------------------------------------------------------------------------------------------
KEY_LINE:
'------------------------------------------------------------------------------------------------------------------------------------
    If KEY_STR = "" Then: GoTo 1983
    LINE_STR = KEY_STR & ":" & ITEM_STR
    Call COLLECT_OBJ.Add(LINE_STR, KEY_STR)
    If Err.number <> 0 Then
        LINE_STR = COLLECT_OBJ.Item(KEY_STR) & "," & ITEM_STR
        Call COLLECT_OBJ.Remove(KEY_STR)
        Call COLLECT_OBJ.Add(LINE_STR, KEY_STR)
        Err.Clear
    End If
1983:
'------------------------------------------------------------------------------------------------------------------------------------
Return
'------------------------------------------------------------------------------------------------------------------------------------
VECTOR_LINE:
'------------------------------------------------------------------------------------------------------------------------------------
    If i <> j Then l = 1 Else l = 1 - j
    TEMP_VECTOR(i + l, 1) = URL_STR & SYMBOL_STR & "&m=" & ITEM_STR 'http://finance.yahoo.com/q/os?s=CSCO&m=2011-04
    TEMP_VECTOR(i + l, 2) = k
    TEMP_VECTOR(i + l, 3) = ITEM_STR
'------------------------------------------------------------------------------------------------------------------------------------
Return
'------------------------------------------------------------------------------------------------------------------------------------
MATRIX_LINE:
'------------------------------------------------------------------------------------------------------------------------------------
    NROWS = COLLECT_OBJ.COUNT
    NCOLUMNS = j
    ReDim TEMP_MATRIX(0 To NROWS, 1 To NCOLUMNS + 1)
    TEMP_MATRIX(0, 1) = SYMBOL_STR
    For k = 1 To NCOLUMNS
        If k <> NCOLUMNS Then l = 2 Else l = 2 - NCOLUMNS
        TEMP_MATRIX(0, k + l) = Replace(Mid(DATA_STR, h(k), 9), "m=", "")
    Next k
    For k = 1 To NROWS
        For l = 1 To NCOLUMNS: TEMP_MATRIX(k, l + 1) = "": Next l
        LINE_STR = COLLECT_OBJ(k) & ","
        i = 1: j = InStr(i, LINE_STR, ":")
        TEMP_MATRIX(k, 1) = Mid(LINE_STR, i, j - i)
        i = j + 1
        Do While i < Len(LINE_STR)
            j = InStr(i, LINE_STR, ",")
            If j = 0 Then: Exit Do
            For l = 1 To NCOLUMNS
                If TEMP_MATRIX(0, l + 1) = Mid(LINE_STR, i, j - i) Then: TEMP_MATRIX(k, l + 1) = True '"X"
            Next l
            i = j + 1
        Loop
    Next k
    TEMP_MATRIX = MATRIX_QUICK_SORT_FUNC(TEMP_MATRIX, 1, 0)
'------------------------------------------------------------------------------------------------------------------------------------
Return
'------------------------------------------------------------------------------------------------------------------------------------

ERROR_LABEL:
YAHOO_OPTIONS_EXPIRY_STRIKE_FUNC = Err.number
End Function


Function YAHOO_OPTION_QUOTES_FUNC(ByVal TICKER_STR As String, _
Optional ByVal STRIKE_MONTH As String = "", _
Optional ByVal RESET_FLAG As Boolean = False)

'STRIKE_MONTH (e.g., 2011-01):
'Possible input field reserved to allow use of option months
'other than the nearest one.

Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim ii As Long
Dim jj As Long
Dim kk As Long
Dim ll As Long

Dim NROWS As Long

Dim KEY_STR As String
Dim REF_STR As String
Dim TEMP_STR As String
Dim DATA_STR As String
Dim SRC_URL_STR As String

Dim SPOT_STR As String
Dim EXPIRATION_STR As String

Dim SRC_URL_ARR() As String
Dim TEMP1_ARR() As String
Dim TEMP2_ARR As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

'----------------------------------------------------------------------------------
Call HASH_YAHOO_OPTION_QUOTES_FUNC(RESET_FLAG)
KEY_STR = TICKER_STR & "|" & STRIKE_MONTH
'----------------------------------------------------------------------------------
If PUB_YAHOO_OPTIONS_HASH_OBJ.Exists(KEY_STR) Then
    TEMP_MATRIX = PUB_YAHOO_OPTIONS_HASH_OBJ.Item(KEY_STR)
    YAHOO_OPTION_QUOTES_FUNC = CONVERT_STRING_NUMBER_FUNC(TEMP_MATRIX)
    Exit Function
End If
'----------------------------------------------------------------------------------
SRC_URL_STR = "http://finance.yahoo.com/q/os?s=" & TICKER_STR
'----------------------------------------------------------------------------------
If STRIKE_MONTH = "" Then 'Data for all Expiration Dates
'----------------------------------------------------------------------------------
    DATA_STR = SAVE_WEB_DATA_PAGE_FUNC(SRC_URL_STR, 0, False, 0, False)
    
    ii = InStr(1, DATA_STR, "time_rtq_ticker")
    If ii = 0 Then: GoTo ERROR_LABEL
    ii = InStr(ii, DATA_STR, ">")
    ii = ii + 1
    
    jj = InStr(ii, DATA_STR, ">")
    kk = InStr(jj, DATA_STR, "<")
    
    SPOT_STR = Mid(DATA_STR, jj + 1, kk - jj - 1)
    
    k = 1
    kk = 1
    ReDim SRC_URL_ARR(1 To k)
    SRC_URL_ARR(k) = SRC_URL_STR
    kk = InStr(1 + kk, DATA_STR, "m=2")
    k = k + 1
    Do While kk <> 0
        ReDim Preserve SRC_URL_ARR(1 To k)
        SRC_URL_ARR(k) = SRC_URL_ARR(1) & "&" & Mid(DATA_STR, kk, 9)
        kk = InStr(1 + kk, DATA_STR, "m=2")
        k = k + 1
    Loop
    k = k - 2
    ReDim Preserve SRC_URL_ARR(1 To k)
'----------------------------------------------------------------------------------
Else
'----------------------------------------------------------------------------------
    ReDim SRC_URL_ARR(1 To 1)
    SRC_URL_STR = SRC_URL_STR & CStr("&m=") & STRIKE_MONTH
    SRC_URL_ARR(1) = SRC_URL_STR
    
    DATA_STR = SAVE_WEB_DATA_PAGE_FUNC(SRC_URL_STR, 0, False, 0, False)
    
    ii = InStr(1, DATA_STR, "time_rtq_ticker")
    If ii = 0 Then: GoTo ERROR_LABEL
    ii = InStr(ii, DATA_STR, ">")
    ii = ii + 1
    
    jj = InStr(ii, DATA_STR, ">")
    kk = InStr(jj, DATA_STR, "<")
    
    SPOT_STR = Mid(DATA_STR, jj + 1, kk - jj - 1)
'----------------------------------------------------------------------------------
End If
'----------------------------------------------------------------------------------
NROWS = 1
ReDim TEMP1_ARR(1 To NROWS) 'Load Data For Each Expiration
'----------------------------------------------------------------------------------
For k = LBound(SRC_URL_ARR, 1) To UBound(SRC_URL_ARR, 1)
'----------------------------------------------------------------------------------
   SRC_URL_STR = SRC_URL_ARR(k)
   
   DATA_STR = SAVE_WEB_DATA_PAGE_FUNC(SRC_URL_STR, 0, False, 0, False)
   DATA_STR = Replace(DATA_STR, Chr(10), "")
   DATA_STR = Replace(DATA_STR, "</span>", "")
   DATA_STR = Replace(DATA_STR, "</b>", "")
   
   REF_STR = "Options Expiring"
   ii = InStr(1, DATA_STR, REF_STR)
   If ii = 0 Then: GoTo 1983
   ii = ii + Len(REF_STR)
   
   jj = InStr(ii, DATA_STR, "<")
   If jj = 0 Then: GoTo 1983
   EXPIRATION_STR = Trim(Mid(DATA_STR, ii, jj - ii))
   
   REF_STR = "</tr></table></td>"
   jj = InStr(ii, DATA_STR, REF_STR)
   If jj = 0 Then: GoTo 1983
   
   DATA_STR = Mid(DATA_STR, ii, jj - ii)
   
   kk = 1
   REF_STR = "/q?s="
   ii = InStr(1, DATA_STR, REF_STR)
'--------------------------------------------------------------------
   Do While ii <> 0
'--------------------------------------------------------------------
       TEMP_STR = EXPIRATION_STR
'--------------------------------------------------------------------
       For kk = 1 To 2 'Calls & Puts
'--------------------------------------------------------------------
           ii = ii + Len(REF_STR)
           ii = InStr(ii, DATA_STR, ">")
           If ii = 0 Then: GoTo 1982
           ii = ii + 1
           jj = InStr(ii, DATA_STR, "<")
           If jj = 0 Then: GoTo 1982
           TEMP_STR = TEMP_STR & "|" & Trim(Mid(DATA_STR, ii, jj - ii)) 'Symbol
'--------------------------------------------------------------------
           For ll = 1 To 6
                ii = jj
                ii = ii + 10
                jj = InStr(ii, DATA_STR, "</td>")
                If jj = 0 Then: GoTo 1982
                           
                ii = jj
                Do While Mid(DATA_STR, ii, 1) <> ">"
                     ii = ii - 1
                Loop
                ii = ii + 1
                TEMP_STR = TEMP_STR & "|" & Replace(Trim(Mid(DATA_STR, ii, jj - ii)), "N/A", 0)
                'Last/Change/Bid/Ask/Volume/Open Int
           Next ll
'--------------------------------------------------------------------
           If kk = 1 Then
                ii = InStr(jj, DATA_STR, "k=")
                If ii = 0 Then: GoTo 1982
                ii = InStr(ii, DATA_STR, ">")
                If ii = 0 Then: GoTo 1982
                ii = ii + 1
                jj = InStr(ii, DATA_STR, "<")
                If jj = 0 Then: GoTo 1982
                TEMP_STR = TEMP_STR & "|" & Trim(Mid(DATA_STR, ii, jj - ii)) 'Strike
           End If
'--------------------------------------------------------------------
1982:
           ii = InStr(jj, DATA_STR, REF_STR)
           If ii = 0 Then: Exit For
'--------------------------------------------------------------------
       Next kk
'--------------------------------------------------------------------

       ReDim Preserve TEMP1_ARR(1 To NROWS)
       TEMP_STR = TEMP_STR & "|"
       TEMP1_ARR(NROWS) = TEMP_STR
       NROWS = NROWS + 1
'--------------------------------------------------------------------
   Loop
'--------------------------------------------------------------------

1983:
'----------------------------------------------------------------------------------
Next k
'----------------------------------------------------------------------------------
NROWS = NROWS - 1
ReDim TEMP_MATRIX(0 To NROWS, 1 To 17)
DATA_STR = "SPOT,EXPIRING,CALLS,LAST,CHANGE,BID,ASK,VOLUME,OPEN INT,STRIKE,PUTS,LAST,CHANGE,BID,ASK,VOLUME,OPEN INT,"
l = Len(DATA_STR)
i = 1: k = 1
Do While i <= l
    j = InStr(i, DATA_STR, ",")
    TEMP_MATRIX(0, k) = Mid(DATA_STR, i, j - i)
    i = j + 1
    k = k + 1
Loop
For i = 1 To NROWS
    DATA_STR = TEMP1_ARR(i)
    TEMP_MATRIX(i, 1) = SPOT_STR
    ii = 1: jj = InStr(ii, DATA_STR, "|")
    TEMP_STR = Mid(DATA_STR, ii, jj - ii)
    TEMP2_ARR = Split(TEMP_STR, ",")
    TEMP_STR = TEMP2_ARR(LBound(TEMP2_ARR) + 1)
    If Len(TEMP2_ARR(UBound(TEMP2_ARR))) = 2 Then
        TEMP_STR = TEMP_STR & " " & "20" & TEMP2_ARR(UBound(TEMP2_ARR))
    Else
        TEMP_STR = TEMP_STR & " " & TEMP2_ARR(UBound(TEMP2_ARR))
    End If
    TEMP_MATRIX(i, 2) = DateValue(TEMP_STR)
    ii = jj + 1
    For j = 3 To 17
        jj = InStr(ii, DATA_STR, "|")
        If jj = 0 Then: Exit For
        TEMP_MATRIX(i, j) = Mid(DATA_STR, ii, jj - ii)
        ii = jj + 1
    Next j
Next i

PUB_YAHOO_OPTIONS_HASH_OBJ.Add KEY_STR, TEMP_MATRIX
YAHOO_OPTION_QUOTES_FUNC = CONVERT_STRING_NUMBER_FUNC(TEMP_MATRIX)

Exit Function
ERROR_LABEL:
YAHOO_OPTION_QUOTES_FUNC = Err.number
End Function


Function YAHOO_OPTIONS_PARSE_TICKERS_FUNC(ByRef TICKERS_RNG As Variant)

Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim ii As Long
Dim jj As Long
Dim kk As Long

Const CHR_STR As String = "00000000"
Dim TICKER_STR As String

Dim TEMP_VAL As Variant
Dim NROWS As Long
Dim TICKERS_VECTOR As Variant
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
NROWS = UBound(TICKERS_VECTOR, 1)

ReDim TEMP_MATRIX(0 To NROWS, 1 To 4)
TEMP_MATRIX(0, 1) = "SYMBOL"
TEMP_MATRIX(0, 2) = "TYPE"
TEMP_MATRIX(0, 3) = "EXPIRATION"
TEMP_MATRIX(0, 4) = "STRIKE"

For k = 1 To NROWS
    TICKER_STR = TICKERS_VECTOR(k, 1) 'GOOG110408C00555000
    If TICKER_STR = "" Then: GoTo 1983
    l = Len(TICKER_STR): j = l
    l = Len(CHR_STR): i = j - l + 1
    If i <= 0 Then: GoTo 1983
    TEMP_VAL = CLng(Mid(TICKER_STR, i, j - i + 1))
    If TEMP_VAL = 0 Then: GoTo 1983
    TEMP_VAL = TEMP_VAL / 1000 'Format(1000 * STRIKE0_VAL, "00000000")
    TEMP_MATRIX(k, 4) = TEMP_VAL
    j = i: i = j - 1 'PC: PC
    TEMP_MATRIX(k, 2) = Mid(TICKER_STR, i, j - i)
    j = i: i = j - 6 'Expiration: Format(EXPIRY0_VAL, "yymmdd")
    TEMP_VAL = Mid(TICKER_STR, i, j - i)
    j = i: i = 1 'Ticker
    TEMP_MATRIX(k, 1) = Mid(TICKER_STR, i, j - i)
    TEMP_VAL = "20" & TEMP_VAL
    i = 1: j = 4: ii = Mid(TEMP_VAL, i, j)
    i = 5: j = 2: jj = Mid(TEMP_VAL, i, j)
    i = 7: j = 2: kk = Mid(TEMP_VAL, i, j)
    TEMP_MATRIX(k, 3) = DateSerial(ii, jj, kk)
1983:
Next k
YAHOO_OPTIONS_PARSE_TICKERS_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
YAHOO_OPTIONS_PARSE_TICKERS_FUNC = Err.number
End Function

Private Sub HASH_YAHOO_OPTION_QUOTES_FUNC( _
Optional ByVal RESET_FLAG As Boolean = False)

On Error Resume Next

'--------------------------------------------------------------------
If RESET_FLAG = False Then
'--------------------------------------------------------------------
    If PUB_YAHOO_OPTIONS_HASH_OBJ Is Nothing Then
        Set PUB_YAHOO_OPTIONS_HASH_OBJ = New clsTypeHash
        PUB_YAHOO_OPTIONS_HASH_OBJ.SetSize 10000
        PUB_YAHOO_OPTIONS_HASH_OBJ.IgnoreCase = False
    End If
'--------------------------------------------------------------------
Else
'--------------------------------------------------------------------
    Set PUB_YAHOO_OPTIONS_HASH_OBJ = New clsTypeHash
    PUB_YAHOO_OPTIONS_HASH_OBJ.SetSize 10000
    PUB_YAHOO_OPTIONS_HASH_OBJ.IgnoreCase = False
'--------------------------------------------------------------------
End If
'--------------------------------------------------------------------
End Sub

Sub PRINT_YAHOO_OPTION_QUOTES_FUNC()

Dim SROW As Long
Dim SCOLUMN As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TICKER_STR As String
Dim DST_RNG As Excel.Range
Dim TEMP_RNG As Excel.Range
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

TICKER_STR = Excel.Application.InputBox("Symbol", "Yahoo Finance")
Call EXCEL_TURN_OFF_EVENTS_FUNC
        
Set DST_RNG = WSHEET_ADD_FUNC(PARSE_CURRENT_TIME_FUNC("_"), _
              ActiveWorkbook).Cells(3, 3)
        
TEMP_MATRIX = YAHOO_OPTION_QUOTES_FUNC(TICKER_STR, "", False)
If IsArray(TEMP_MATRIX) = False Then: GoTo 1983
        
SROW = LBound(TEMP_MATRIX, 1)
NROWS = UBound(TEMP_MATRIX, 1)
            
SCOLUMN = LBound(TEMP_MATRIX, 2)
NCOLUMNS = UBound(TEMP_MATRIX, 2)
            
Set TEMP_RNG = Range(DST_RNG.Cells(SROW, SCOLUMN), _
DST_RNG.Cells(NROWS, NCOLUMNS))
TEMP_RNG.value = TEMP_MATRIX
GoSub FORMAT_LINE

1983:
Call EXCEL_TURN_ON_EVENTS_FUNC

Exit Sub
'-----------------------------------------------------------------------------
FORMAT_LINE:
'-----------------------------------------------------------------------------
    With TEMP_RNG
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .WEIGHT = xlThin
        End With
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .WEIGHT = xlThin
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .WEIGHT = xlThin
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .WEIGHT = xlThin
        End With
        With .Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .WEIGHT = xlThin
        End With
        With .Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .WEIGHT = xlThin
        End With
        With .Rows(1)
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
            With .Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 65535
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        End With
        .ColumnWidth = 15
        .RowHeight = 15
    End With
    Return
'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------
ERROR_LABEL:
Call EXCEL_TURN_ON_EVENTS_FUNC
End Sub
