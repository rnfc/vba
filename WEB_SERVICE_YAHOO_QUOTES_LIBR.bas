Attribute VB_Name = "WEB_SERVICE_YAHOO_QUOTES_LIBR"

'Option Compare Text  'Uppercase letters to be equivalent to lowercase letters.
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

Public PUB_YAHOO_QUOTES_CONNECT_URL_STR As String
Private PUB_YAHOO_QUOTES_CONNECT_OBJ As MSXML2.XMLHTTP60

Private PUB_YAHOO_QUOTES_CODES_OBJ As Collection
Public Const PUB_YAHOO_QUOTES_CODES_STR As String = _
"m6,% off 200-day avg,m8,% off 50-day avg,t8,1yr target price,m4,200-day moving avg,m3,50-day moving avg," & _
"k,52-week high,j,52-week low,w,52-week range,c8,after hours change (ecn),g3,annualized gain,a,ask,b2,ask (ecn)," & _
"a5,ask size,a2,average daily volume,b,bid,b3,bid (ecn),b6,bid size,b4,book value,c1,change,c,change & percent," & _
"k2,change & percent (ecn),c6,change (ecn),m5,change from 200-day moving avg,m7,change from 50-day moving avg," & _
"k4,change from 52-week high,j5,change from 52-week low,c3,commission,d1,date of last trade,m,day's range," & _
"m2,Day's range (ecn),w1,day's value change,w4,day's value change (ecn),r1,dividend pay date,y,dividend yield," & _
"d,dividend/share,e,earnings/share,j4,ebitda,e7,eps est. current yr,e9,eps est. next quarter,e8,eps est. next year," & _
"e1,error indication (returned for symbol changed / invalid),x,exchange,q,ex-dividend date,e3,expiration date," & _
"f6,float shares,h,high,l2,high limit,g4,holdings gain,g5,holdings gain & percent (ecn),g6,holdings gain (ecn)," & _
"g1,holdings gain percent,v1,holdings value,v7,holdings value (ecn),l1,last trade,k1,last trade (ecn with time)," & _
"l,last trade (with time),k3,last trade size,g,low,l3,low limit,j3,market cap (ecn),j1,market capitalization," & _
"i,more info,n,name,n4,notes,o,open,o1,open interest?,i5,order book (ecn),r2,p/e (ecn),r,p/e ratio," & _
"k5,pct chg from 52-week high,j6,pct chg from 52-week low,r5,peg ratio,p2,percent change,p,previous close," & _
"p1,price paid,p6,price/book,r6,price/eps est. current yr,r7,price/eps est. next yr,p5,price/sales,s6,revenue," & _
"s1,shares owned,s7,short ratio,s3,strike price,s,symbol,t7,ticker trend,t1,time of last trade,d2,trade date," & _
"t6,trade links,p3,type of option,v,volume,"


Function YAHOO_QUOTES_FUNC(ByRef TICKERS_RNG As Variant, _
ByRef ELEMENTS_RNG As Variant, _
Optional ByVal REFRESH_CALLER As Variant = 0, _
Optional ByVal HEADER_FLAG As Boolean = True, _
Optional ByVal SERVER_STR As String = "")

Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim l As Integer

Dim ii As Integer
Dim jj As Integer
Dim kk As Integer
Dim ll As Integer

Dim MAX1_VAL As Integer
Dim MAX2_VAL As Integer

Dim NROWS As Integer
Dim NCOLUMNS As Integer

Dim DELIM_CHR As String

Dim ELEMENT_STR As String
Dim ELEMENTS_STR As String

Dim TICKER_STR As String
Dim TICKERS_STR As String

Dim ELEMENTS_VECTOR As Variant
Dim TICKERS_VECTOR As Variant

Dim DATA_VECTOR As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DELIM_CHR = YAHOO_QUOTES_SERVER_DELIM_FUNC(SERVER_STR)
'-------------------------------------------------------------------------------
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
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
If IsArray(ELEMENTS_RNG) = True Then
    ELEMENTS_VECTOR = ELEMENTS_RNG
    If UBound(ELEMENTS_VECTOR, 2) = 1 Then
        ELEMENTS_VECTOR = MATRIX_TRANSPOSE_FUNC(ELEMENTS_VECTOR)
    End If
Else
    ReDim ELEMENTS_VECTOR(1 To 1, 1 To 1)
    ELEMENTS_VECTOR(1, 1) = ELEMENTS_RNG
End If
NCOLUMNS = UBound(ELEMENTS_VECTOR, 2)
'-------------------------------------------------------------------------------

MAX1_VAL = 200
MAX2_VAL = 50
'-------------------------------------------------------------------------------
If NROWS <= MAX1_VAL And NCOLUMNS <= MAX2_VAL Then
'-------------------------------------------------------------------------------
    For k = 1 To NROWS
        TICKER_STR = Trim(CStr(TICKERS_VECTOR(k, 1)))
        If TICKER_STR = "" Or TICKER_STR = "0" Then: TICKER_STR = "--"
        If k < NROWS Then
            TICKERS_STR = TICKERS_STR & TICKERS_VECTOR(k, 1) & DELIM_CHR
        Else
            TICKERS_STR = TICKERS_STR & TICKER_STR
        End If
    Next k
    For l = 1 To NCOLUMNS
        ELEMENT_STR = Trim(CStr(ELEMENTS_VECTOR(1, l)))
        If ELEMENT_STR = "" Or ELEMENT_STR = "0" Then: ELEMENT_STR = "Symbol"
        If l < NCOLUMNS Then
            ELEMENTS_STR = ELEMENTS_STR & YAHOO_QUOTES_CODES_DESCRIPTION_FUNC(ELEMENT_STR) & DELIM_CHR
        Else
            ELEMENTS_STR = ELEMENTS_STR & YAHOO_QUOTES_CODES_DESCRIPTION_FUNC(ELEMENT_STR)
        End If
    Next l
    DATA_MATRIX = MATRIX_YAHOO_QUOTES_FUNC(TICKERS_STR, ELEMENTS_STR, SERVER_STR, REFRESH_CALLER, HEADER_FLAG, "+", 0, 0)
'-------------------------------------------------------------------------------
Else
'-------------------------------------------------------------------------------
    
    If HEADER_FLAG = True Then
        ReDim DATA_MATRIX(0 To NROWS, 0 To NCOLUMNS)
        DATA_MATRIX(0, 0) = "symbol"
        For i = 1 To NROWS: DATA_MATRIX(i, 0) = TICKERS_VECTOR(i, 1): Next i
        For j = 1 To NCOLUMNS: DATA_MATRIX(0, j) = ELEMENTS_VECTOR(1, j): Next j
    Else
        ReDim DATA_MATRIX(1 To NROWS, 1 To NCOLUMNS)
    End If
    ii = 1: kk = 0
    Do While ii <= NROWS
        TICKERS_STR = ""
        For k = 1 To MAX1_VAL
            TICKER_STR = Trim(CStr(TICKERS_VECTOR(ii, 1)))
            If TICKER_STR = "" Or TICKER_STR = "0" Then: TICKER_STR = "--"
            If ii < NROWS Then
                TICKERS_STR = TICKERS_STR & TICKER_STR & DELIM_CHR
                ii = ii + 1
            Else
                TICKERS_STR = TICKERS_STR & TICKER_STR
                ii = ii + 1
                Exit For
            End If
        Next k
        
        jj = 1
        ll = 0
        Do While jj <= NCOLUMNS
            ELEMENTS_STR = ""
            For l = 1 To MAX2_VAL
                ELEMENT_STR = Trim(CStr(ELEMENTS_VECTOR(1, jj)))
                If ELEMENT_STR = "" Or ELEMENT_STR = "0" Then: ELEMENT_STR = "Symbol"
                If jj < NCOLUMNS Then
                    ELEMENTS_STR = ELEMENTS_STR & YAHOO_QUOTES_CODES_DESCRIPTION_FUNC(ELEMENT_STR) & DELIM_CHR
                    jj = jj + 1
                Else
                    ELEMENTS_STR = ELEMENTS_STR & YAHOO_QUOTES_CODES_DESCRIPTION_FUNC(ELEMENT_STR)
                    jj = jj + 1
                    Exit For
                End If
            Next l
            DATA_VECTOR = MATRIX_YAHOO_QUOTES_FUNC(TICKERS_STR, ELEMENTS_STR, SERVER_STR, REFRESH_CALLER, False, "+", 0, 0)
            If IsArray(DATA_VECTOR) = False Then: Exit Do
            For i = 1 To UBound(DATA_VECTOR, 1)
                For j = 1 To UBound(DATA_VECTOR, 2)
                    DATA_MATRIX(kk + i, ll + j) = DATA_VECTOR(i, j)
                Next j
            Next i
            ll = ll + UBound(DATA_VECTOR, 2) - 1
        Loop
        kk = kk + UBound(DATA_VECTOR, 1) - 1
    Loop
1983:
End If
'-------------------------------------------------------------------------------

YAHOO_QUOTES_FUNC = DATA_MATRIX

Exit Function
ERROR_LABEL:
YAHOO_QUOTES_FUNC = Err.number
End Function

'Returns delayed stock quotes and other data from the Yahoo quotes interface.

Function MATRIX_YAHOO_QUOTES_FUNC(ByVal TICKERS_STR As String, _
ByRef ELEMENTS_STR As String, _
Optional ByVal SERVER_STR As String = "", _
Optional ByVal REFRESH_CALLER As Variant = 0, _
Optional ByVal HEADER_FLAG As Boolean = False, _
Optional ByRef PREFIX_STR As String = "+", _
Optional ByRef NROWS As Integer = 0, _
Optional ByRef NCOLUMNS As Integer = 0)

Dim DELIM_CHR As String
Dim DATA_STR As String

Dim TEMP1_STR As String
Dim TEMP2_STR As String

On Error GoTo ERROR_LABEL

DELIM_CHR = YAHOO_QUOTES_SERVER_DELIM_FUNC(SERVER_STR)

TEMP1_STR = ELEMENTS_STR
If NCOLUMNS = 0 Then: _
NCOLUMNS = COUNT_CHARACTERS_FUNC(TEMP1_STR, DELIM_CHR) + 1 'Columns

TEMP1_STR = Replace(TEMP1_STR, " ", "")
TEMP1_STR = Replace(TEMP1_STR, DELIM_CHR, "")

TEMP2_STR = TICKERS_STR
If NROWS = 0 Then: NROWS = COUNT_CHARACTERS_FUNC(TEMP2_STR, DELIM_CHR) + 1 ' Rows

TEMP2_STR = Replace(TEMP2_STR, DELIM_CHR, PREFIX_STR)

If HEADER_FLAG = True Then
    ReDim DATA_MATRIX(0 To NROWS, 0 To NCOLUMNS)
    DATA_MATRIX = YAHOO_QUOTES_HEADER_FUNC(DATA_MATRIX, TICKERS_STR, ELEMENTS_STR, DELIM_CHR)
Else
    ReDim DATA_MATRIX(1 To NROWS, 1 To NCOLUMNS)
End If
PUB_YAHOO_QUOTES_CONNECT_URL_STR = YAHOO_QUOTES_URL_FUNC(SERVER_STR, TEMP1_STR, TEMP2_STR, PREFIX_STR)


'Call REMOVE_WEB_PAGE_CACHE_FUNC(PUB_YAHOO_QUOTES_CONNECT_URL_STR)
'DATA_STR = XML_HTTP_SYNCHRONOUS_FUNC(PUB_YAHOO_QUOTES_CONNECT_URL_STR, "GET", "") & Chr(13)
GoSub CONNECT_LINE
DATA_MATRIX = PARSE_YAHOO_QUOTES_FUNC(DATA_MATRIX, DATA_STR, DELIM_CHR)
MATRIX_YAHOO_QUOTES_FUNC = DATA_MATRIX

'In most cases, this will need to be an array-entered formula. To
'array-enter a formula in EXCEL, first highlight the range of cells
'where you would like the returned data to appear -- the number of
'rows for the range should be AT LEAST the number of ticker symbols
'you are requesting from the function, while the number of columns
'for the range should be AT LEAST the number of data items you are
'requesting for each ticker symbol from the function. Next, enter
'your formula and then press Ctrl-Shift-Enter.

'-------------------------------------------------------------------------------------------------------------------
Exit Function
'-------------------------------------------------------------------------------------------------------------------
CONNECT_LINE:
'-------------------------------------------------------------------------------------------------------------------
Call REMOVE_WEB_PAGE_CACHE_FUNC(PUB_YAHOO_QUOTES_CONNECT_URL_STR)
If PUB_YAHOO_QUOTES_CONNECT_OBJ Is Nothing Then: Set PUB_YAHOO_QUOTES_CONNECT_OBJ = New MSXML2.XMLHTTP60
With PUB_YAHOO_QUOTES_CONNECT_OBJ
    .Open "GET", PUB_YAHOO_QUOTES_CONNECT_URL_STR, False
    .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    .send ""
    If .Status = 200 Then DATA_STR = .ResponseText & Chr(13)
End With
'-------------------------------------------------------------------------------------------------------------------
Return
'-------------------------------------------------------------------------------------------------------------------
ERROR_LABEL:
MATRIX_YAHOO_QUOTES_FUNC = Err.number
End Function

Function YAHOO_QUOTES_HEADER_FUNC(ByRef DATA_RNG As Variant, _
ByVal TICKERS_STR As String, _
ByVal ELEMENTS_STR As String, _
Optional ByVal DELIM_CHR As String = ",")

Dim i As Integer
Dim j As Integer
Dim k As Integer

Dim NROWS As Integer
Dim NCOLUMNS As Integer

On Error GoTo ERROR_LABEL

NROWS = UBound(DATA_RNG, 1)
NCOLUMNS = UBound(DATA_RNG, 2)

DATA_RNG(0, 0) = "symbol"

i = 1
j = i
For k = 1 To NROWS
    i = InStr(i, TICKERS_STR, DELIM_CHR)
    If i = 0 Then
        i = Len(TICKERS_STR) + 1
        DATA_RNG(k, 0) = Mid(TICKERS_STR, j, i - j)
        Exit For
    Else
        DATA_RNG(k, 0) = Mid(TICKERS_STR, j, i - j)
    End If
    i = i + 1
    j = i
Next k

i = 1
j = i
For k = 1 To NCOLUMNS
    i = InStr(i, ELEMENTS_STR, DELIM_CHR)
    If i = 0 Then
        i = Len(ELEMENTS_STR) + 1
        DATA_RNG(0, k) = YAHOO_QUOTES_CODES_DESCRIPTION_FUNC(Mid(ELEMENTS_STR, j, i - j))
        Exit For
    Else
        DATA_RNG(0, k) = YAHOO_QUOTES_CODES_DESCRIPTION_FUNC(Mid(ELEMENTS_STR, j, i - j))
    End If
    i = i + 1
    j = i
Next k

YAHOO_QUOTES_HEADER_FUNC = DATA_RNG

Exit Function
ERROR_LABEL:
YAHOO_QUOTES_HEADER_FUNC = Err.number
End Function

'Parse returned data

Function PARSE_YAHOO_QUOTES_FUNC(ByRef DATA_RNG As Variant, _
ByRef DATA_STR As String, _
ByVal DELIM_CHR As String)

Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim l As Integer
'Dim m As Integer

Dim ii As Integer
Dim jj As Integer

Dim NROWS As Integer
Dim NCOLUMNS As Integer

Dim TEMP1_STR As String
Dim TEMP2_STR As String
Dim TEMP3_STR As String
Dim TEMP4_STR As String

Dim LINE_STR As String
'Dim TEMP1_VAL As Variant
'Dim TEMP2_VAL As Variant
Dim DATA_ARR As Variant

On Error GoTo ERROR_LABEL

NROWS = UBound(DATA_RNG, 1)
NCOLUMNS = UBound(DATA_RNG, 2)

GoSub HOUSE_KEEPING
DATA_ARR = Split(DATA_STR, vbLf)

l = LBound(DATA_ARR, 1)
For i = 1 To NROWS
    If i > (UBound(DATA_ARR) - LBound(DATA_ARR) + 1) Then: Exit For
    ii = 1
    LINE_STR = DATA_ARR(l)
    GoSub CLEAN_LINE
    For j = 1 To NCOLUMNS
        TEMP1_STR = LINE_STR 'Missing Thousand Separator
        If ii > Len(TEMP1_STR) Then Exit For
        TEMP2_STR = IIf(Mid(TEMP1_STR, ii, 1) = Chr(34), Chr(34), "") & DELIM_CHR
        jj = InStr(ii, TEMP1_STR & DELIM_CHR, TEMP2_STR)
        
        If ii = 1 Then
            If Mid(TEMP1_STR, jj + 1, Len(" *.")) Like " *." Then: jj = InStr(jj + 1, TEMP1_STR & DELIM_CHR, TEMP2_STR)
            'Name in the first entry of the array with DelimChr and without Chr(34)
        End If
        '----------------------------------------------------------------------------------------------------
        'If (jj - ii - Len(TEMP2_STR) + 1) <= 0 Or (ii + Len(TEMP2_STR) - 1) <= 0 Then: Exit For
        '255 Text Limit in a Range --> QueryTable Destination is a Range
        TEMP1_STR = Left(Mid(TEMP1_STR, ii + Len(TEMP2_STR) - 1, jj - ii - Len(TEMP2_STR) + 1), 255)
        TEMP1_STR = Trim(TEMP1_STR)
        DATA_RNG(i, j) = TEMP1_STR
        '----------------------------------------------------------------------------------------------------
        'TEMP1_VAL = Left(Mid(TEMP1_STR, ii + Len(TEMP2_STR) - 1, jj - ii - Len(TEMP2_STR) + 1), 255)
        'TEMP2_VAL = Trim(TEMP1_VAL)
        'If Right(TEMP2_VAL, 1) = "%" Then
        '   m = 100
        '   TEMP2_VAL = Left(TEMP2_VAL, Len(TEMP2_VAL) - 1)
        'Else
        '   m = 1
        'End If
        'On Error Resume Next
        'TEMP1_VAL = CDec(TEMP2_VAL) / m
        'On Error GoTo ERROR_LABEL
        'DATA_RNG(i, j) = TEMP1_VAL
        '----------------------------------------------------------------------------------------------------
        ii = jj + Len(TEMP2_STR)
    Next j
    l = l + 1
Next i

PARSE_YAHOO_QUOTES_FUNC = CONVERT_STRING_NUMBER_FUNC(DATA_RNG)

Exit Function
'-----------------------------------------------------------------------------------------------
HOUSE_KEEPING:
'-----------------------------------------------------------------------------------------------
'DATA_STR = Replace(DATA_STR, Chr(10), Chr(13))
'DATA_STR = Replace(DATA_STR, Chr(13) & Chr(13), Chr(13))
'DATA_STR = Replace(DATA_STR, Chr(13) & Chr(13), Chr(13))
'Debug.Print DATA_STR

DATA_STR = Replace(DATA_STR, vbCrLf, vbLf)
DATA_STR = Replace(DATA_STR, Chr(13), "")
'DATA_STR = Trim(DATA_STR)
DATA_STR = REMOVE_EXTRA_SPACES_FUNC(DATA_STR)
DATA_STR = Replace(DATA_STR, "," & Space(1), ",")
'DATA_STR = Replace(DATA_STR, Chr(34) & "," & Space(1), Chr(34) & ",")

DATA_STR = Replace(DATA_STR, "</b>", "")
DATA_STR = Replace(DATA_STR, "<b>", "")

'DATA_STR = Replace(DATA_STR, "&nbsp;-&nbsp;", "")
'DATA_STR = Replace(DATA_STR, "&nbsp;+&nbsp;", "")
'DATA_STR = Replace(DATA_STR, "&nbsp;+", "")
'DATA_STR = Replace(DATA_STR, "&nbsp;-", "")
DATA_STR = Replace(DATA_STR, "&nbsp;", "")

DATA_STR = Replace(DATA_STR, "-=X", "")
DATA_STR = Replace(DATA_STR, "</i>", "")
DATA_STR = Replace(DATA_STR, "<i>", "")
DATA_STR = Replace(DATA_STR, "N/A", Chr(34) & "0" & Chr(34))

'YAHOO_QUOTES_REPLACE_BOLD_TAGS_FUNC
'OUT_STR = Trim(Replace(IN_STR, "<b>", ""))
'OUT_STR = Trim(Replace(OUT_STR, "</b>", ""))
'OUT_STR = Trim(Replace(OUT_STR, "<i>", ""))
'OUT_STR = Trim(Replace(OUT_STR, "</i>", ""))
'OUT_STR = Trim(Replace(OUT_STR, "&nbsp;", ""))

Return
'-----------------------------------------------------------------------------------------------
CLEAN_LINE: 'Remove Thousands Separator for FLoating Shares/Bid Size/Ask Size....
'-----------------------------------------------------------------------------------------------
    If Left(LINE_STR, 1) <> DELIM_CHR Then: LINE_STR = DELIM_CHR & LINE_STR
    If Right(LINE_STR, 1) <> DELIM_CHR Then: LINE_STR = LINE_STR & DELIM_CHR
    k = 1
    jj = Len(LINE_STR)
    'LINE_STR = "324.44,3,400,45.43,23,345,45.33,23,456,23.45,567,783"
    '-----------------------------------------------------------------------------------
    Do Until k = 0
    '-------------------------------------------------------------------------------------
    'Or Mid(LINE_STR, k, 16) Like vbLf & "##,###,###,###" & DELIM_CHR Or _
        Mid(LINE_STR, k, 16) Like DELIM_CHR & "##,###,###,###" & vbLf
        If Mid(LINE_STR, k, 16) Like DELIM_CHR & "##,###,###,###" & DELIM_CHR Then
            TEMP1_STR = Mid(LINE_STR, 1, k + 2)
            TEMP2_STR = Mid(LINE_STR, k + 4, 3)
            TEMP3_STR = Mid(LINE_STR, k + 8, 3)
            TEMP4_STR = Mid(LINE_STR, k + 12, jj)
            LINE_STR = TEMP1_STR & TEMP2_STR & TEMP3_STR & TEMP4_STR
            k = k + 12
        End If
        'Debug.Print Mid(LINE_STR, k, jj)
        If Mid(LINE_STR, k, 15) Like DELIM_CHR & "#,###,###,###" & DELIM_CHR Then
            TEMP1_STR = Mid(LINE_STR, 1, k + 1)
            TEMP2_STR = Mid(LINE_STR, k + 3, 3)
            TEMP3_STR = Mid(LINE_STR, k + 7, 3)
            TEMP4_STR = Mid(LINE_STR, k + 11, jj)
            LINE_STR = TEMP1_STR & TEMP2_STR & TEMP3_STR & TEMP4_STR
            k = k + 11
        End If
        'Debug.Print Mid(LINE_STR, k, jj)
    '-------------------------------------------------------------------------------------
        If Mid(LINE_STR, k, 13) Like DELIM_CHR & "###,###,###" & DELIM_CHR Then
            TEMP1_STR = Mid(LINE_STR, 1, k + 3)
            TEMP2_STR = Mid(LINE_STR, k + 5, 3)
            TEMP3_STR = Mid(LINE_STR, k + 9, jj)
            LINE_STR = TEMP1_STR & TEMP2_STR & TEMP3_STR
            k = k + 10
        End If
        'Debug.Print Mid(LINE_STR, k, jj)
        If Mid(LINE_STR, k, 12) Like DELIM_CHR & "##,###,###" & DELIM_CHR Then
            TEMP1_STR = Mid(LINE_STR, 1, k + 2)
            TEMP2_STR = Mid(LINE_STR, k + 4, 3)
            TEMP3_STR = Mid(LINE_STR, k + 8, jj)
            LINE_STR = TEMP1_STR & TEMP2_STR & TEMP3_STR
            k = k + 9
        End If
        'Debug.Print Mid(LINE_STR, k, jj)
        If Mid(LINE_STR, k, 11) Like DELIM_CHR & "#,###,###" & DELIM_CHR Then
            TEMP1_STR = Mid(LINE_STR, 1, k + 1)
            TEMP2_STR = Mid(LINE_STR, k + 3, 3)
            TEMP3_STR = Mid(LINE_STR, k + 7, jj)
            LINE_STR = TEMP1_STR & TEMP2_STR & TEMP3_STR
            k = k + 8
        End If
        'Debug.Print Mid(LINE_STR, k, jj)
    '-------------------------------------------------------------------------------------
        If Mid(LINE_STR, k, 9) Like DELIM_CHR & "###,###" & DELIM_CHR Then
            TEMP1_STR = Mid(LINE_STR, 1, k + 3)
            TEMP2_STR = Mid(LINE_STR, k + 5, jj)
            LINE_STR = TEMP1_STR & TEMP2_STR
            k = k + 7
        End If
        'Debug.Print Mid(LINE_STR, k, jj)
        If Mid(LINE_STR, k, 8) Like DELIM_CHR & "##,###" & DELIM_CHR Then
            TEMP1_STR = Mid(LINE_STR, 1, k + 2)
            TEMP2_STR = Mid(LINE_STR, k + 4, jj)
            LINE_STR = TEMP1_STR & TEMP2_STR
            k = k + 6
        End If
        'Debug.Print Mid(LINE_STR, k, jj)
        If Mid(LINE_STR, k, 7) Like DELIM_CHR & "#,###" & DELIM_CHR Then
            TEMP1_STR = Mid(LINE_STR, 1, k + 1)
            TEMP2_STR = Mid(LINE_STR, k + 3, jj)
            LINE_STR = TEMP1_STR & TEMP2_STR
            k = k + 5
        End If
        'Debug.Print Mid(LINE_STR, k, jj)
        k = InStr(k + 1, LINE_STR, DELIM_CHR)
    '-----------------------------------------------------------------------------------
    Loop
    '-----------------------------------------------------------------------------------
    If Left(LINE_STR, 1) = DELIM_CHR Then
        LINE_STR = Mid(LINE_STR, 2, jj)
        jj = jj - 1
    End If
    
    If Right(LINE_STR, 1) = DELIM_CHR Then
        LINE_STR = Mid(LINE_STR, 1, jj - 1)
    End If
'-----------------------------------------------------------------------------------------------
Return
'-----------------------------------------------------------------------------------------------
ERROR_LABEL:
PARSE_YAHOO_QUOTES_FUNC = Err.number
End Function

Function YAHOO_QUOTES_URL_FUNC(ByRef SERVER_STR As String, _ByRef ELEMENTS_STR As String, _ByRef TICKERS_STR As String, _Optional ByVal DELIM_CHR As String = "+")Dim SRC_URL_STR As StringOn Error GoTo ERROR_LABEL'Select Case UCase(SERVER_STR)'Case "", "UNITED STATES"     SRC_URL_STR = _        "http://download.finance.yahoo.com/d/quotes.csv?s="        'http://download.finance.yahoo.com/d/quotes.csvr?'Case "JAPAN" '    SRC_URL_STR = "http://finance.yahoo.com." & YAHOO_QUOTES_SERVER_ID_FUNC(SERVER_STR) & "/d/quotes.csv?s="'Case "MEXICO" '    SRC_URL_STR = "http://" & YAHOO_QUOTES_SERVER_ID_FUNC(SERVER_STR) & ".finance.yahoo.com/d/quotes.csv?s="  '   TICKERS_STR = Replace(TICKERS_STR, DELIM_CHR, YAHOO_QUOTES_SERVER_DELIM_FUNC(SERVER_STR))'Case Else'     SRC_URL_STR = "http://" & YAHOO_QUOTES_SERVER_ID_FUNC(SERVER_STR) & ".finance.yahoo.com/d/quotes.csv?s="'End Select'SERVERURL=http://sg.finance.yahoo.com'STOCKS=OCBC.SI+JARD.SI+0307.HK'Format = snl1d1t1c1p2pomwverdybaxq'#SERVERURL=http://finanzen.de.yahoo.com'#STOCKS=185775.F'##FORMAT=snl1d1t1c1p2poghwerdy'#FORMAT=snl1d1t1c1p2poghwv'#a2werr1dyj1'cd /tmp'rm -vf quote*csv*'wget -q "$SERVERURL/d/quotes.csv?f=$FORMAT&s=$STOCKS"'cat quote*csv*' http://download.finance.yahoo.com/d/quotes.csv?s=%40%5EDJI,GOOG&f=nsl1op&e=.csvSRC_URL_STR = SRC_URL_STR & TICKERS_STR & "&f=" & LCase(ELEMENTS_STR) & "&e=.csv" '& "?" & Now()'Debug.Print SRC_URL_STRYAHOO_QUOTES_URL_FUNC = SRC_URL_STRExit FunctionERROR_LABEL:YAHOO_QUOTES_URL_FUNC = Err.numberEnd Function

Function YAHOO_QUOTES_SERVER_ID_FUNC(ByRef SERVER_STR As Variant)
    
On Error GoTo ERROR_LABEL

Select Case UCase(SERVER_STR)
Case "ARGENTINA": YAHOO_QUOTES_SERVER_ID_FUNC = "ar"
Case "AUSTRALIA": YAHOO_QUOTES_SERVER_ID_FUNC = "au"
Case "BRAZIL": YAHOO_QUOTES_SERVER_ID_FUNC = "br"
Case "CANADA": YAHOO_QUOTES_SERVER_ID_FUNC = "ca"
Case "CHINESE": YAHOO_QUOTES_SERVER_ID_FUNC = "chinese"
Case "CHINA": YAHOO_QUOTES_SERVER_ID_FUNC = "cn"
Case "DENMARK": YAHOO_QUOTES_SERVER_ID_FUNC = "de"
Case "FRANCE": YAHOO_QUOTES_SERVER_ID_FUNC = "fr"
Case "FRENCH CANADA": YAHOO_QUOTES_SERVER_ID_FUNC = "cf"
Case "GERMANY": YAHOO_QUOTES_SERVER_ID_FUNC = "de"
Case "HONG KONG": YAHOO_QUOTES_SERVER_ID_FUNC = "hk"
Case "INDIA": YAHOO_QUOTES_SERVER_ID_FUNC = "in"
Case "ITALY": YAHOO_QUOTES_SERVER_ID_FUNC = "it"
Case "JAPAN": YAHOO_QUOTES_SERVER_ID_FUNC = "jp"
Case "KOREA": YAHOO_QUOTES_SERVER_ID_FUNC = "kr"
Case "MEXICO": YAHOO_QUOTES_SERVER_ID_FUNC = "mx"
Case "SINGAPORE": YAHOO_QUOTES_SERVER_ID_FUNC = "sg"
Case "SPAIN": YAHOO_QUOTES_SERVER_ID_FUNC = "es"
Case "UNITED KINGDOM": YAHOO_QUOTES_SERVER_ID_FUNC = "uk"
End Select

Exit Function
ERROR_LABEL:
YAHOO_QUOTES_SERVER_ID_FUNC = Err.number
End Function


Function YAHOO_QUOTES_SERVER_DELIM_FUNC(Optional ByVal SERVER_STR As String)
    
On Error GoTo ERROR_LABEL

Select Case UCase(SERVER_STR)
Case "ARGENTINA": YAHOO_QUOTES_SERVER_DELIM_FUNC = ";"
Case "AUSTRALIA": YAHOO_QUOTES_SERVER_DELIM_FUNC = ","
Case "BRAZIL": YAHOO_QUOTES_SERVER_DELIM_FUNC = ","
Case "CANADA": YAHOO_QUOTES_SERVER_DELIM_FUNC = ","
Case "CHINA": YAHOO_QUOTES_SERVER_DELIM_FUNC = ","
Case "FRANCE": YAHOO_QUOTES_SERVER_DELIM_FUNC = ";"
Case "DENMARK": YAHOO_QUOTES_SERVER_DELIM_FUNC = ";"
Case "HONG KONG": YAHOO_QUOTES_SERVER_DELIM_FUNC = ","
Case "INDIA": YAHOO_QUOTES_SERVER_DELIM_FUNC = ","
Case "ITALY": YAHOO_QUOTES_SERVER_DELIM_FUNC = ";"
Case "JAPAN": YAHOO_QUOTES_SERVER_DELIM_FUNC = ","
Case "KOREA": YAHOO_QUOTES_SERVER_DELIM_FUNC = ","
Case "MEXICO": YAHOO_QUOTES_SERVER_DELIM_FUNC = ","
Case "SINGAPORE": YAHOO_QUOTES_SERVER_DELIM_FUNC = ","
Case "SPAIN": YAHOO_QUOTES_SERVER_DELIM_FUNC = ";"
Case "UNITED KINGDOM": YAHOO_QUOTES_SERVER_DELIM_FUNC = ","
'Case "UNITED STATES": YAHOO_QUOTES_SERVER_DELIM_FUNC = ","
Case Else: YAHOO_QUOTES_SERVER_DELIM_FUNC = "," 'US'
End Select

Exit Function
ERROR_LABEL:
YAHOO_QUOTES_SERVER_DELIM_FUNC = Err.number
End Function

Function YAHOO_QUOTES_CODES_DESCRIPTION_FUNC(ByVal CODE_STR As String)

Dim i As Long
Dim j As Long
Dim k As Long

Dim KEY_STR As String
Dim ITEM_STR As String

On Error GoTo ERROR_LABEL

If PUB_YAHOO_QUOTES_CODES_OBJ Is Nothing Then
    Set PUB_YAHOO_QUOTES_CODES_OBJ = New Collection
    k = Len(PUB_YAHOO_QUOTES_CODES_STR): i = 1
    Do
        j = InStr(i, PUB_YAHOO_QUOTES_CODES_STR, ",")
        If j = 0 Then: GoTo ERROR_LABEL
        KEY_STR = Mid(PUB_YAHOO_QUOTES_CODES_STR, i, j - i)
        i = j + 1
        j = InStr(i, PUB_YAHOO_QUOTES_CODES_STR, ",")
        If j = 0 Then: GoTo ERROR_LABEL
        ITEM_STR = Mid(PUB_YAHOO_QUOTES_CODES_STR, i, j - i)
        i = j + 1
        Call PUB_YAHOO_QUOTES_CODES_OBJ.Add(ITEM_STR, KEY_STR) 'KEY_STR, KEY_STR)
        Call PUB_YAHOO_QUOTES_CODES_OBJ.Add(KEY_STR, ITEM_STR)
    Loop Until i > k
End If

On Error Resume Next
ITEM_STR = PUB_YAHOO_QUOTES_CODES_OBJ.Item(CODE_STR)
If Err.number <> 0 Then
    Err.Clear
    YAHOO_QUOTES_CODES_DESCRIPTION_FUNC = "--"
Else
    YAHOO_QUOTES_CODES_DESCRIPTION_FUNC = ITEM_STR
End If

Exit Function
ERROR_LABEL:
YAHOO_QUOTES_CODES_DESCRIPTION_FUNC = "--"
End Function

Function YAHOO_QUOTES_TICKER_TREND_COUNT_FUNC(ByVal TREND_STR As String) As Integer

Dim i As Long
Dim j As Long
Dim k As Long

On Error GoTo ERROR_LABEL

k = 0
j = Len(TREND_STR)    'Length of String
For i = 1 To j          'Increment thru
    Select Case Mid(TREND_STR, i, 1)
        Case "+"                'If it is a comma
            k = k + 1
        Case "-"
            k = k - 1
        Case "="
            k = k
    End Select
Next i

YAHOO_QUOTES_TICKER_TREND_COUNT_FUNC = k

Exit Function
ERROR_LABEL:
YAHOO_QUOTES_TICKER_TREND_COUNT_FUNC = 0
End Function

Function YAHOO_QUOTES_INVESTMENT_STYLE_FUNC(Optional ByVal MARKET_CAP_VAL As Double = 0)

Dim OUT_STR As String

On Error GoTo ERROR_LABEL

If ((MARKET_CAP_VAL > 0) And (MARKET_CAP_VAL <= 50000000)) Then
    OUT_STR = "Nano-Cap"
ElseIf ((MARKET_CAP_VAL > 50000000) And (MARKET_CAP_VAL <= 300000000)) Then
    OUT_STR = "Micro-Cap"
ElseIf ((MARKET_CAP_VAL > 300000000) And (MARKET_CAP_VAL <= 2000000000)) Then
    OUT_STR = "Small-Cap"
ElseIf ((MARKET_CAP_VAL > 2000000000) And (MARKET_CAP_VAL <= 10000000000#)) Then
    OUT_STR = "Mid-Cap"
ElseIf (MARKET_CAP_VAL > 10000000000#) Then
    OUT_STR = "Large-Cap"
Else
    OUT_STR = "Undefined"
End If

YAHOO_QUOTES_INVESTMENT_STYLE_FUNC = OUT_STR

Exit Function
ERROR_LABEL:
YAHOO_QUOTES_INVESTMENT_STYLE_FUNC = ""
End Function

Function YAHOO_QUOTES_STRING_TO_NUMBER_FUNC(ByVal IN_STR As String) As Double

Dim OUT_STR As String
Dim TEMP_VAL As Double

On Error GoTo ERROR_LABEL

OUT_STR = IN_STR

If IsNumeric(IN_STR) Then
    YAHOO_QUOTES_STRING_TO_NUMBER_FUNC = IN_STR
Else
    If InStr(1, OUT_STR, "B", vbTextCompare) Then
        OUT_STR = Replace(OUT_STR, "B", "")
        TEMP_VAL = OUT_STR * 1000000000
    ElseIf InStr(1, OUT_STR, "M", vbTextCompare) Then
        OUT_STR = Replace(OUT_STR, "M", "")
        TEMP_VAL = OUT_STR * 1000000
    ElseIf InStr(1, OUT_STR, "K", vbTextCompare) Then
        OUT_STR = Replace(OUT_STR, "K", "")
        TEMP_VAL = OUT_STR * 1000
    End If
End If

YAHOO_QUOTES_STRING_TO_NUMBER_FUNC = TEMP_VAL
Exit Function
ERROR_LABEL:
YAHOO_QUOTES_STRING_TO_NUMBER_FUNC = 0
End Function
