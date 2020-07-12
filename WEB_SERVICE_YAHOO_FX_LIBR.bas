Attribute VB_Name = "WEB_SERVICE_YAHOO_FX_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

Public PUB_YAHOO_FX_CODES_STR As String
Private PUB_YAHOO_FX_CODES_OBJ As Collection


'************************************************************************************
'************************************************************************************
'FUNCTION      : YAHOO_FX_QUOTES_FUNC

'DESCRIPTION   : This module gets FX rates from Yahoo! Finance.  The algo will
'return a quote for each of the FX symbols passed to it.  It also allows to
'specify a format other than the default to take advantage of the extended range
'of available information.

'The download operation is efficient: only one request is made even if
'several symbols are requested at once. The return value is an array,
'with the following elements: Bid, Ask, and last traded price.

'LIBRARY       : YAHOO
'GROUP         : WEB SERVICE
'ID            : 001
'LAST UPDATE   : 12/05/2011
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function YAHOO_FX_QUOTES_FUNC(ByVal BASE_FX_RNG As Variant, _
ByVal QUOTE_FX_RNG As Variant, _
Optional ByVal REFRESH_CALLER As Variant = 0, _
Optional ByVal HEADER_FLAG As Boolean = False)
'http://www.federalreserve.gov/releases/H10/
'If REFRESH_CALLER <> 0 Then: Excel.Application.Volatile (True)

Dim i As Integer
Dim NROWS As Integer

Dim TICKERS_STR As String
Dim ELEMENTS_STR As String

Dim BASE_VECTOR As Variant
Dim QUOTE_VECTOR As Variant

Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

If IsArray(BASE_FX_RNG) = True Then
    BASE_VECTOR = BASE_FX_RNG
    If UBound(BASE_VECTOR) = 1 Then: _
        BASE_VECTOR = MATRIX_TRANSPOSE_FUNC(BASE_VECTOR)
Else
    ReDim BASE_VECTOR(1 To 1, 1 To 1)
    BASE_VECTOR(1, 1) = BASE_FX_RNG
End If

NROWS = UBound(BASE_VECTOR, 1)

If IsArray(QUOTE_FX_RNG) = True Then
    QUOTE_VECTOR = QUOTE_FX_RNG
    If UBound(QUOTE_VECTOR) = 1 Then: QUOTE_VECTOR = MATRIX_TRANSPOSE_FUNC(QUOTE_VECTOR)
    If UBound(QUOTE_VECTOR, 1) <> NROWS Then: GoTo ERROR_LABEL
Else
    ReDim QUOTE_VECTOR(1 To NROWS, 1 To 1)
    For i = 1 To NROWS: QUOTE_VECTOR(i, 1) = QUOTE_FX_RNG: Next i
End If

TICKERS_STR = ""
For i = 1 To NROWS
    If (QUOTE_VECTOR(i, 1) <> "") And (BASE_VECTOR(i, 1) <> "") Then
        TICKERS_STR = TICKERS_STR & YAHOO_FX_CODES_DESCRIPTION_FUNC(CStr(BASE_VECTOR(i, 1))) & YAHOO_FX_CODES_DESCRIPTION_FUNC(CStr(QUOTE_VECTOR(i, 1))) & "=X" & ","
    Else
        TICKERS_STR = TICKERS_STR & "--" & ","
    End If
Next i

TICKERS_STR = Left(TICKERS_STR, Len(TICKERS_STR) - 1)
ELEMENTS_STR = "b,a,l1,t1,d1"

TEMP_MATRIX = MATRIX_YAHOO_QUOTES_FUNC(TICKERS_STR, ELEMENTS_STR, "", REFRESH_CALLER, HEADER_FLAG, "+", 0, 0)
YAHOO_FX_QUOTES_FUNC = TEMP_MATRIX
    
Exit Function
ERROR_LABEL:
YAHOO_FX_QUOTES_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : YAHOO_FX_CROSS_MATRIX_FUNC
'DESCRIPTION   : Displays an exchange rate table for cross currency rates
'LIBRARY       : YAHOO
'GROUP         : WEB SERVICE
'ID            : 002
'LAST UPDATE   : 12/05/2011
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function YAHOO_FX_CROSS_MATRIX_FUNC(Optional ByRef TICKERS_RNG As Variant, _
Optional ByVal REFRESH_CALLER As Variant = 0, _
Optional ByVal HEADER_FLAG As Boolean = True, _
Optional ByVal OUTPUT As Integer = 1)

'TICKERS --> USD, CAD, EUR, JPY, GBP, CHF....
'FX Cross Rate Matrix --> $1 is worth... 0.6770 Euro: USDEUR=X

Dim i As Integer
Dim j As Integer
Dim k As Integer

Dim ii As Integer
Dim jj As Integer

Dim NSIZE As Integer

Dim TICKERS_STR As String
'Dim ELEMENT_VECTOR As Variant
Dim TICKERS_VECTOR As Variant

Dim TEMP_VECTOR As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

'--------------------------------------------------------------------------------------------
If IsArray(TICKERS_RNG) = True Then
'--------------------------------------------------------------------------------------------
    TICKERS_VECTOR = TICKERS_RNG
    If UBound(TICKERS_VECTOR, 1) = 1 Then
        TICKERS_VECTOR = MATRIX_TRANSPOSE_FUNC(TICKERS_VECTOR)
    End If
'--------------------------------------------------------------------------------------------
Else
'--------------------------------------------------------------------------------------------
    TICKERS_STR = "USD,EUR,JPY,GBP,CHF,CAD,AUD,HKD,"
    j = Len(TICKERS_STR)
    k = 0
    For i = 1 To j
        If Mid(TICKERS_STR, i, 1) = "," Then: k = k + 1
    Next i
    ReDim TICKERS_VECTOR(1 To k, 1 To 1)
    ii = 1
    For i = 1 To k
        jj = InStr(ii, TICKERS_STR, ",")
        TICKERS_VECTOR(i, 1) = Mid(TICKERS_STR, ii, jj - ii)
        ii = jj + 1
    Next i
'--------------------------------------------------------------------------------------------
End If
'--------------------------------------------------------------------------------------------
NSIZE = UBound(TICKERS_VECTOR, 1)
If HEADER_FLAG = True Then
    ReDim TEMP_MATRIX(0 To NSIZE, 0 To NSIZE)
Else
    ReDim TEMP_MATRIX(1 To NSIZE, 1 To NSIZE)
End If
'--------------------------------------------------------------------------------------------
'ReDim TEMP_VECTOR(1 To (NSIZE * NSIZE) - NSIZE + 1, 1 To 1)
TICKERS_STR = ""
jj = 1
For j = 1 To NSIZE
    For i = 1 To NSIZE
        If i <> j Then
            TICKERS_STR = TICKERS_STR & TICKERS_VECTOR(i, 1) & TICKERS_VECTOR(j, 1) & "=X" & ","
            'TEMP_VECTOR(jj, 1) = TICKERS_VECTOR(i, 1) & TICKERS_VECTOR(j, 1) & "=X"
            jj = jj + 1
        End If
    Next i
Next j
TICKERS_STR = Left(TICKERS_STR, Len(TICKERS_STR) - 1)

'--------------------------------------------------------------------------------------------
Select Case OUTPUT
'--------------------------------------------------------------------------------------------
Case 0
'--------------------------------------------------------------------------------------------
    TEMP_VECTOR = MATRIX_YAHOO_QUOTES_FUNC(TICKERS_STR, "l1", "", REFRESH_CALLER, False, "+", 0, 0) 'Last Trade --> l1
    'TEMP_VECTOR = YAHOO_QUOTES_FUNC(TEMP_VECTOR, "Last Trade", REFRESH_CALLER, False, "")
    If HEADER_FLAG = True Then
        TEMP_MATRIX(0, 0) = "--"
        For j = 1 To NSIZE
            TEMP_MATRIX(0, j) = TICKERS_VECTOR(j, 1)
            TEMP_MATRIX(j, 0) = TICKERS_VECTOR(j, 1)
        Next j
    End If
    jj = 0
    For j = 1 To NSIZE
        ii = 1
        For i = 1 To NSIZE
            If i <> j Then 'Using Last Traded Price
                TEMP_MATRIX(i, j) = TEMP_VECTOR(ii + jj, 1)
                ii = ii + 1
            Else
                TEMP_MATRIX(i, j) = 1
            End If
        Next i
        jj = jj + NSIZE - 1
    Next j
'--------------------------------------------------------------------------------------------
Case Else 'For the FX Arbitrage Algorithm
'--------------------------------------------------------------------------------------------
'    ReDim ELEMENT_VECTOR(1 To 2, 1 To 1)
'    ELEMENT_VECTOR(1, 1) = "Ask"
'    ELEMENT_VECTOR(2, 1) = "Bid"
    TEMP_VECTOR = MATRIX_YAHOO_QUOTES_FUNC(TICKERS_STR, "a,b", "", REFRESH_CALLER, False, "+", 0, 0) 'Last Trade --> l1
    'TEMP_VECTOR = YAHOO_QUOTES_FUNC(TEMP_VECTOR, ELEMENT_VECTOR, REFRESH_CALLER, False, "")
    
    If HEADER_FLAG = True Then
        TEMP_MATRIX(0, 0) = "--"
        For j = 1 To NSIZE
            TEMP_MATRIX(0, j) = TICKERS_VECTOR(j, 1)
            TEMP_MATRIX(j, 0) = TICKERS_VECTOR(NSIZE - j + 1, 1)
        Next j
    End If
    k = 1
    jj = 0
    For i = NSIZE To 1 Step -1
        ii = 1
        For j = 1 To NSIZE
            If j < k Then 'Ask
                TEMP_MATRIX(i, j) = TEMP_VECTOR(ii + jj, 1)
                ii = ii + 1
            ElseIf j = k Then
                TEMP_MATRIX(i, j) = "--"
            Else 'Bid
                TEMP_MATRIX(i, j) = TEMP_VECTOR(ii + jj, 2)
                ii = ii + 1
            End If
        Next j
        k = k + 1
        jj = jj + NSIZE - 1
    Next i
'--------------------------------------------------------------------------------------------
End Select
'--------------------------------------------------------------------------------------------

YAHOO_FX_CROSS_MATRIX_FUNC = TEMP_MATRIX
    
Exit Function
ERROR_LABEL:
YAHOO_FX_CROSS_MATRIX_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : YAHOO_FX_CODES_DESCRIPTION_FUNC
'DESCRIPTION   : World Currency Symbols
'LIBRARY       : YAHOO
'GROUP         : WEB SERVICE
'ID            : 003
'LAST UPDATE   : 12/05/2011
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function YAHOO_FX_CODES_DESCRIPTION_FUNC(ByVal CODE_STR As String)

'Debug.Print YAHOO_FX_CODES_DESCRIPTION_FUNC("British Pound (GBP)")
'Debug.Print YAHOO_FX_CODES_DESCRIPTION_FUNC("GBP")
'Debug.Print "-----------------------------------------------------------------------------"
'Debug.Print PUB_YAHOO_FX_CODES_STR

Dim i As Long
Dim j As Long
Dim k As Long

Dim KEY_STR As String
Dim DATA_STR As String
Dim ITEM_STR As String
Dim ELEMENTS_STR As String
Dim SRC_URL_STR As String
On Error GoTo ERROR_LABEL

If PUB_YAHOO_FX_CODES_OBJ Is Nothing Then
    SRC_URL_STR = "http://finance.yahoo.com/currency-converter/"
    DATA_STR = SAVE_WEB_DATA_PAGE_FUNC(SRC_URL_STR, 0, False, 0, False)
    KEY_STR = ">Type a country or currency"
    i = InStr(1, DATA_STR, KEY_STR): i = i + Len(KEY_STR)
    
    KEY_STR = "<div": j = InStr(i, DATA_STR, KEY_STR)
    DATA_STR = Mid(DATA_STR, i, j - i)
    
    KEY_STR = "value='"
    ELEMENTS_STR = "": i = 1
    i = InStr(i, DATA_STR, KEY_STR)
    PUB_YAHOO_FX_CODES_STR = ""
    Do
        i = i + Len(KEY_STR): j = InStr(i, DATA_STR, "'")
        ITEM_STR = Mid(DATA_STR, i, j - i): ELEMENTS_STR = ELEMENTS_STR & ITEM_STR & ","
    
        i = j + 2: j = InStr(i, DATA_STR, "<")
        ITEM_STR = Mid(DATA_STR, i, j - i): ELEMENTS_STR = ELEMENTS_STR & ITEM_STR & ","
        PUB_YAHOO_FX_CODES_STR = PUB_YAHOO_FX_CODES_STR & ITEM_STR & ","
        
        i = j: i = InStr(i, DATA_STR, KEY_STR)
    Loop Until i = 0
'    Debug.Print ELEMENTS_STR
'    Debug.Print PUB_YAHOO_FX_CODES_STR
    
    Set PUB_YAHOO_FX_CODES_OBJ = New Collection
    k = Len(ELEMENTS_STR): i = 1
    Do
        j = InStr(i, ELEMENTS_STR, ",")
        If j = 0 Then: GoTo ERROR_LABEL
        KEY_STR = Mid(ELEMENTS_STR, i, j - i)
        i = j + 1
        j = InStr(i, ELEMENTS_STR, ",")
        If j = 0 Then: GoTo ERROR_LABEL
        ITEM_STR = Mid(ELEMENTS_STR, i, j - i)
        i = j + 1
        Call PUB_YAHOO_FX_CODES_OBJ.Add(ITEM_STR, KEY_STR)
        Call PUB_YAHOO_FX_CODES_OBJ.Add(KEY_STR, ITEM_STR)
    Loop Until i > k
End If

On Error Resume Next
ITEM_STR = PUB_YAHOO_FX_CODES_OBJ.Item(CODE_STR)
If Err.number <> 0 Then
    Err.Clear
    YAHOO_FX_CODES_DESCRIPTION_FUNC = "--"
Else
    YAHOO_FX_CODES_DESCRIPTION_FUNC = ITEM_STR
End If

Exit Function
ERROR_LABEL:
YAHOO_FX_CODES_DESCRIPTION_FUNC = "--"
End Function
