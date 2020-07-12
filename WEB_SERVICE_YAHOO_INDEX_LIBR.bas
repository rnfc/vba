Attribute VB_Name = "WEB_SERVICE_YAHOO_INDEX_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

Public PUB_YAHOO_INDEX_SUFFIX_STR As String

'************************************************************************************
'************************************************************************************
'FUNCTION      :
'DESCRIPTION   :
'LIBRARY       : YAHOO
'GROUP         :
'ID            : 001
'LAST UPDATE   : 29/02/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function YAHOO_INDEX_QUOTES_FUNC(ByRef TICKER_STR As String, _
ByRef ELEMENTS_RNG As Variant, _
Optional ByVal PAGE_VAL As Variant = "", _
Optional ByVal HEADER_FLAG As Boolean = False, _
Optional ByVal REFRESH_CALLER As Variant = 0, _
Optional ByVal SERVER_STR As String = "")

'TICKER_STR: ^GSPC

Dim h As Integer
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim l As Integer

Dim ii As Integer
Dim jj As Integer
Dim kk As Integer

Dim NROWS As Integer
Dim NCOLUMNS As Integer

Dim DELIM_CHR As String
Dim ELEMENT_STR As String
Dim ELEMENT_VECTOR As Variant

Dim DATA1_MATRIX As Variant
Dim DATA2_MATRIX As Variant

Dim DATA_GROUP As Variant

On Error GoTo ERROR_LABEL

DELIM_CHR = YAHOO_QUOTES_SERVER_DELIM_FUNC(SERVER_STR)

kk = 50
'-------------------------------------------------------------------------------
If IsArray(ELEMENTS_RNG) = False Then
    ReDim ELEMENT_VECTOR(1 To 1, 1 To 1)
    ELEMENT_VECTOR(1, 1) = ELEMENTS_RNG
Else
    ELEMENT_VECTOR = ELEMENTS_RNG
    If UBound(ELEMENT_VECTOR, 2) = 1 Then: _
        ELEMENT_VECTOR = MATRIX_TRANSPOSE_FUNC(ELEMENT_VECTOR)
End If
'-------------------------------------------------------------------------------
NROWS = kk + 1
NCOLUMNS = UBound(ELEMENT_VECTOR, 2)

ELEMENT_STR = ""
For i = 1 To NCOLUMNS - 1
    If (Trim(ELEMENT_VECTOR(1, i)) = "") Or (ELEMENT_VECTOR(1, i) = 0) Then: ELEMENT_VECTOR(1, i) = "Symbol"
    ELEMENT_STR = ELEMENT_STR & YAHOO_QUOTES_CODES_DESCRIPTION_FUNC(CStr(ELEMENT_VECTOR(1, i))) & DELIM_CHR
Next i
ELEMENT_STR = ELEMENT_STR & YAHOO_QUOTES_CODES_DESCRIPTION_FUNC(CStr(ELEMENT_VECTOR(1, NCOLUMNS)))

ELEMENT_STR = "n" & "," & ELEMENT_STR
NCOLUMNS = UBound(ELEMENT_VECTOR, 2) + 1

TICKER_STR = "@" & TICKER_STR
'--------------------------------------------------------------------------------------
If PAGE_VAL <> "" Then
'--------------------------------------------------------------------------------------
    PUB_YAHOO_INDEX_SUFFIX_STR = "&h=" & CStr(PAGE_VAL * kk)
    DATA1_MATRIX = MATRIX_YAHOO_QUOTES_FUNC(TICKER_STR, ELEMENT_STR, SERVER_STR, REFRESH_CALLER, False, "+", NROWS, 0)
    If PAGE_VAL > 0 Then h = 1 Else h = 0
    NROWS = NROWS - 1
'------------------------------------------------------------------------------
    If HEADER_FLAG = True Then
'------------------------------------------------------------------------------
        ReDim DATA2_MATRIX(0 To NROWS, 0 To NCOLUMNS - 1)
        DATA2_MATRIX(0, 0) = "Ticker"
        For i = 1 To NROWS: DATA2_MATRIX(i, 0) = DATA1_MATRIX(h + i, 1): Next i
        For j = 1 To NCOLUMNS - 1: DATA2_MATRIX(0, j) = ELEMENT_VECTOR(1, j): Next j
        For j = 2 To NCOLUMNS
            For i = 1 To NROWS
                DATA2_MATRIX(i, j - 1) = DATA1_MATRIX(h + i, j)
            Next i
        Next j
'------------------------------------------------------------------------------
    Else
'------------------------------------------------------------------------------
        ReDim DATA2_MATRIX(1 To NROWS, 1 To NCOLUMNS)
        For j = 1 To NCOLUMNS
            For i = 1 To NROWS
                DATA2_MATRIX(i, j) = DATA1_MATRIX(h + i, j)
            Next i
        Next j
'------------------------------------------------------------------------------
    End If
'------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------
Else
'--------------------------------------------------------------------------------------
    l = 0: k = 0: i = 0
    ReDim DATA_GROUP(1 To 1)
    ReDim DATA2_MATRIX(1 To NROWS - 1, 1 To NCOLUMNS)
'------------------------------------------------------------------------------
    Do
'------------------------------------------------------------------------------
        PUB_YAHOO_INDEX_SUFFIX_STR = "&h=" & CStr(i * kk)
        DATA1_MATRIX = MATRIX_YAHOO_QUOTES_FUNC(TICKER_STR, ELEMENT_STR, SERVER_STR, REFRESH_CALLER, False, "+", NROWS, 0)
        If i > 0 Then
            If DATA1_MATRIX(LBound(DATA1_MATRIX, 1), LBound(DATA1_MATRIX, 2)) = "" Then: Exit Do
            h = 1
        Else
            If IsArray(DATA1_MATRIX) = False Then: GoTo ERROR_LABEL
            h = 0
        End If
        
        For jj = 1 To NCOLUMNS
            For ii = 1 To NROWS - 1
                DATA2_MATRIX(ii, jj) = DATA1_MATRIX(h + ii, jj)
            Next ii
        Next jj
        
        If i > 0 Then
            If DATA_GROUP(1)(2, 1) = DATA2_MATRIX(1, 1) Then: Exit Do
        End If
        
        l = l + UBound(DATA2_MATRIX, 1)
        k = UBound(DATA2_MATRIX, 2)
        
        i = i + 1
        ReDim Preserve DATA_GROUP(1 To i)
        DATA_GROUP(i) = DATA2_MATRIX
'------------------------------------------------------------------------------
    Loop Until i > 500
'------------------------------------------------------------------------------
    
    ReDim DATA1_MATRIX(1 To l, 1 To k)
    
    l = 1
    For k = 1 To UBound(DATA_GROUP)
        For i = 1 To UBound(DATA_GROUP(k), 1)
            For j = 1 To UBound(DATA1_MATRIX, 2)
                DATA1_MATRIX(l, j) = DATA_GROUP(k)(i, j)
            Next j
            l = l + 1
        Next i
    Next k
        
    NROWS = UBound(DATA1_MATRIX, 1)
    NCOLUMNS = UBound(DATA1_MATRIX, 2)

'------------------------------------------------------------------------------
    If HEADER_FLAG = True Then
'------------------------------------------------------------------------------
        ReDim DATA2_MATRIX(0 To NROWS, 0 To NCOLUMNS - 1)
        DATA2_MATRIX(0, 0) = "Ticker"
        For i = 1 To NROWS: DATA2_MATRIX(i, 0) = DATA1_MATRIX(i, 1): Next i
        For j = 1 To NCOLUMNS - 1: DATA2_MATRIX(0, j) = ELEMENT_VECTOR(1, j): Next j
        For j = 2 To NCOLUMNS
            For i = 1 To NROWS
                DATA2_MATRIX(i, j - 1) = DATA1_MATRIX(i, j)
            Next i
        Next j
'------------------------------------------------------------------------------
    Else
'------------------------------------------------------------------------------
        DATA2_MATRIX = DATA1_MATRIX
'------------------------------------------------------------------------------
    End If
'------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------
End If
'--------------------------------------------------------------------------------------
Erase DATA1_MATRIX: Erase DATA_GROUP
YAHOO_INDEX_QUOTES_FUNC = MATRIX_TRIM_FUNC(DATA2_MATRIX, 1, 0)

Exit Function
ERROR_LABEL:
YAHOO_INDEX_QUOTES_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : YAHOO_INDEXES_TICKERS_FUNC
'DESCRIPTION   :
'LIBRARY       : HTML
'GROUP         :
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 14/03/2008
'************************************************************************************
'************************************************************************************

Function YAHOO_INDEXES_TICKERS_FUNC()

Dim i As Long
Dim j As Long
Dim k As Long

Dim ii As Long
Dim jj As Long
Dim kk As Long

Dim TEMP_STR As String
Const DELIM1_CHR As String = "|"
Const DELIM2_CHR As String = ","

Static TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

If IsArray(TEMP_MATRIX) = True Then: GoTo 1983

TEMP_STR = _
"Dow Jones Averages|30 Industrials|^DJI|,Dow Jones Averages|20 Transportation|^DJT|," & _
"Dow Jones Averages|15 Utilities|^DJU|,Dow Jones Averages|65 Composite|^DJA|," & _
"New York Stock Exchange|Volume in 000's|^TV.N|,New York Stock Exchange|Composite|^NYA|," & _
"New York Stock Exchange|Financials|^NFA|,New York Stock Exchange|Industrials|^NDA|," & _
"New York Stock Exchange|Utilities|^NNA|,New York Stock Exchange|Beta Index|^NHB|," & _
"New York Stock Exchange|Tick|^TIC.N|,New York Stock Exchange|ARMS|^STI.N|," & _
"Nasdaq|Composite|^IXIC|,Nasdaq|Volume in 000's|^TV.O|,Nasdaq|National Market Composite|^IXQ|," & _
"Nasdaq|Nasdaq 100|^NDX|,Nasdaq|Banks|^IXBK|,Nasdaq|Financials|^IXFN|,Nasdaq|Financials 100|^IXF|," & _
"Nasdaq|Industrials|^IXID|,Nasdaq|Insurance|^IXIS|,Nasdaq|Computers|^IXK|,Nasdaq|Transportation|^IXTR|," & _
"Nasdaq|Telecommunications|^IXUT|,Nasdaq|Biotech|^NBI|,Standard and Poor's|500 Index|^GSPC|," & _
"Standard and Poor's|100 Index|^OEX|,Standard and Poor's|400 MidCap|^MID|," & _
"Standard and Poor's|600 SmallCap|^SML|,Other U.S. Indices|AMEX Composite|^XAX|," & _
"Other U.S. Indices|AMEX Internet|^IIX|,Other U.S. Indices|AMEX Networking|^NWX|," & _
"Other U.S. Indices|Indi 500|^NDI|,Other U.S. Indices|ISDEX|^IXY2|," & _
"Other U.S. Indices|Major Market|^XMI|,Other U.S. Indices|Pacific Exchange Technology|^PSE|," & _
"Other U.S. Indices|Philadelphia Semiconductor|^SOXX|,Other U.S. Indices|Russell 1000|^RUI|," & _
"Other U.S. Indices|Russell 2000|^RUT|,Other U.S. Indices|Russell 3000|^RUA|," & _
"Other U.S. Indices|TSC Internet|^DOT|,Other U.S. Indices|Value Line|^VLIC|," & _
"Other U.S. Indices|Wilshire 5000 TOT|^TMW|,Treasury Securities (yield x10)|30-Year Bond|^TYX|," & _
"Treasury Securities (yield x10)|10-Year Note|^TNX|,Treasury Securities (yield x10)|5-Year Note|^FVX|," & _
"Treasury Securities (yield x10)|13-Week Bill|^IRX|,Commodities|Dow Jones Spot|^DJS|," & _
"Commodities|Dow Jones Futures|^DJC|,Commodities|Philadelphia Gold & Silver|^XAU|"

k = 1
kk = Len(TEMP_STR)
For ii = 1 To kk
    If Mid(TEMP_STR, ii, 1) = DELIM2_CHR Then: k = k + 1
Next ii

ReDim TEMP_MATRIX(0 To k, 1 To 3)
TEMP_MATRIX(0, 1) = "CATEGORY"
TEMP_MATRIX(0, 2) = "NAME"
TEMP_MATRIX(0, 3) = "SYMBOL"

i = 1
For ii = 1 To k
    For jj = 1 To 3
        j = InStr(i, TEMP_STR, DELIM1_CHR)
        TEMP_MATRIX(ii, jj) = Mid(TEMP_STR, i, j - i)
        i = j + Len(DELIM1_CHR)
    Next jj
    i = i + 1
Next ii

1983:
YAHOO_INDEXES_TICKERS_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
YAHOO_INDEXES_TICKERS_FUNC = Err.number
End Function
